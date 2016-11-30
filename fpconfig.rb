###
### fpconfig.rb
### Author: John Messenger
### License: Apache 2.0 License
###

# This program can read a FreePBX server configuration from the server and write it to an Excel spreadsheet.
# It can also read the spreadsheet format and write it back into the server.
# At present it can only handle Extensions and Trunks.

require 'rubygems'
require 'mechanize'
require 'yaml'
require 'json'
require 'rubyXL'
require 'slop'

ADMIN = '/admin'
CONFIG = ADMIN + '/config.php'
AJAX = ADMIN + '/ajax.php'

SUPPORTED_TABS = %w(Extensions.dahdi Extensions.sip Extensions.iax2 Trunks.dahdi Trunks.sip Trunks.iax2)

####
# Call Mechanize's get method with an optional referer.
# There must be a proper way to do this, but I couldn't figure it out.
####
def get_with_ref(agent, url, params, referer = nil)
  if referer
    agent.get(url, params, referer)
  else
    agent.get(url, params)
  end
end

####
# Call Mechanize's get method with an optional referer, and if FreePBX returns a page with a login form,
# fill it in, submit it and then try again.  Also, if the server is unconfigured, set up the initial credentials
# and skip the free trial offer.
####
def get_page(agent, creds, url, params, referer = nil)
  page = get_with_ref(agent, url, params, referer)
  if /text\/html/ =~ page.response['content-type']
    if page.form && page.form.field_with(value: 'setup_admin')
      setup_admin = page.form
      setup_admin.username = creds[:username]
      setup_admin.password1 = creds[:password]
      setup_admin.password2 = creds[:password]
      abort "Error: can't configure an unconfigured server unless an email address is given with --email or in secrets.yml" unless creds[:email]
      setup_admin.email = creds[:email]
      page = setup_admin.submit
    end
    unless page.css('div#login_form').empty?
      login_form = page.forms[0]
      login_form.username = creds[:username]
      login_form.password = creds[:password]
      page = agent.submit(login_form)
      abort 'Error: login failed' unless page.css('div#login_form').empty?
      page = get_with_ref(agent, url, params, referer)
    end
    if page.form_with(id: 'oobepost')
      skip_trial = page.form
      page = skip_trial.submit(skip_trial.button_with(name: 'skiptrial'))
    end
  end
  page
end

####
# Extract the data from a form.  Construct a table row from it, filtering the data to remove blacklisted keys.
####
def construct_row_from_form(form, my_blacklist, override = {})
  data = {}
  form.fields.each { |ff| data[ff.name] = ff.value }
  form.checkboxes.each { |ff| data[ff.name] = ff.value }
  form.radiobuttons.each do |ff|
    if ff.checked
      data[ff.name] = ff.value
    elsif ff.respond_to? :id
      data[ff.name] = ff.value if override[ff.id]
    end
  end

  my_blacklist.each { |bad| data.delete(bad) }
  data
end

####
# Write the table to a worksheet, ensuring that the set of columns is the superset of the set of keys in all the rows.
####
def write_table(ws, table, my_order)
  keylist = []
  table.each { |row| keylist |= row.keys } # form the superset of all the keys
  colnum = 0
  # First, write the columns for the prioritised keys
  my_order.each do |key|
    ws.add_cell(0, colnum, key.to_s) # the column title
    rownum = 1
    table.each do |row|
      ws.add_cell(rownum, colnum, row[key])
      #row.delete(key)
      rownum += 1
    end
    colnum += 1
    keylist.delete(key)
  end
  # Then, write the columns for the remaining keys
  keylist.each do |key|
    ws.add_cell(0, colnum, key.to_s) # the column title
    rownum = 1
    table.each do |row|
      ws.add_cell(rownum, colnum, row[key])
      #row.delete(key)
      rownum += 1
    end
    colnum += 1
  end
end

####
# Log in to the FreePBX server and read various parameters.  Save them into an Excel spreadsheet.
####
def read_server_write_file(agent, creds, url, outfilename, field_blacklist, field_order, categories)
  wb = RubyXL::Workbook.new

  # Admin worksheet
  ws_admin = wb.add_worksheet('Admin')
  ws_admin.add_cell(0, 0, 'Admin user')
  ws_admin.add_cell(0, 1, 'Password')
  ws_admin.add_cell(1, 0, creds[:username])
  ws_admin.add_cell(1, 1, creds[:password])

  SUPPORTED_TABS.each do |tab|
    category = tab.downcase
    next unless categories.empty? || categories.include?(category.sub(/\..*/, ''))
    ws = nil # Delay creating the worksheet until we know whether there are any entries to put on it

    case category
      # Extensions have a subcategory: tech. Different technologies have different attributes, which means they can't
      # share a tab in the Excel file so easily.
      when /extensions\./
        category,tech = category.split('.')
        ext_page = get_page(agent, creds, url + CONFIG, {display: category})
        ext_grid_result = get_page(agent, creds, url + AJAX,
                                   {module: :core, command: :getExtensionGrid, type: tech, order: :asc}, ext_page.uri.to_s)
        ext_grid = JSON.parse(ext_grid_result.body)
        extn_table = []
        ext_grid.each do |ext|
          ext_page = get_page(agent, creds, url + CONFIG,
                              {display: :extensions, extdisplay: ext['extension']}, ext_grid_result.uri.to_s)
          puts "Extension #{ext['extension']}" unless $quiet
          ws ||= wb.add_worksheet(tab)
          extn_table << construct_row_from_form(ext_page.form('frm_extensions'), field_blacklist[category])
        end
        write_table(ws, extn_table, field_order[category]) if ws

      # Trunks have a subcategory: tech.  Different technologies have different attributes, which means they can't
      # share a tab in the Excel file so easily.
      when /trunks\./
        category,tech = category.split('.')
        trunks_page = get_page(agent, creds, url + CONFIG, {display: :trunks})
        trunks_grid_result = get_page(agent, creds, url + AJAX,
                                      {module: :core, command: :getJSON, jdata: :allTrunks, order: :asc}, trunks_page.uri.to_s)
        trunks_grid = JSON.parse(trunks_grid_result.body)

        trnk_table = []
        trunks_page.css('table#table-all/tbody/tr').each do |tr|  # For each trunk, find its table row...
          next unless tr.css('td')[1].text == tech                # ...ignore if wrong tech
          a = tr.css('td/a')                                      # Find the 'edit' link
          if a and a.first
            linkaddr = a.first['href']                            # Follow trunk's 'edit' link
            trunk_page = get_page(agent, creds, url + ADMIN + '/' + linkaddr, '', trunks_page.uri.to_s)
            # Look up this trunk in the JSON data returned from the AJAX request (without assuming row ordering)
            trk_data = trunks_grid.detect { |e| e['trunkid'] == tr['id'] }
            puts "Trunk #{trk_data['name']}" unless $quiet
            # Tell the form-reader to override certain radioboxes, based on the JSON data.
            # This is necessary because the returned form doesn't contain the currently selected data. Instead,
            # a JavaScript script retrieves some JSON data and sets some of the radioboxes.  It even overrides
            # the value of the 'outcid' field (which is not emulated here).
            override = {
                hcidyes: (/hidden/ =~ trk_data['outcid']),
                hcidno: !(/hidden/ =~ trk_data['outcid']),
                keepcidoff: trk_data['keepcid'] == 'off',
                keepcidon: trk_data['keepcid'] == 'on',
                keepcidcnum: trk_data['keepcid'] == 'cnum',
                keepcidall: trk_data['keepcid'] == 'all',
                continueno: trk_data['continue'] == 'off',
                continueyes: !(trk_data['continue'] == 'off'),
                disabletrunkno: trk_data['disabled'] == 'off',
                disabletrunkyes: !(trk_data['disabled'] == 'off')
            }

            ws ||= wb.add_worksheet(tab)
            trnk_table << construct_row_from_form(trunk_page.form('trunkEdit'), field_blacklist[category], override)
          end
          write_table(ws, trnk_table, field_order[category]) if ws
        end
    end
  end

  #wb.delete_worksheet('Sheet1')
  wb.write(outfilename)
end

####
# Retrieve the form relating to a particular spreadsheet row from the server and fill it in with the provided data.
# Submit the form.
####
def fill_form_and_submit(agent, creds, url, category, tech, data)
  case category
    when 'extensions'
      puts "uploading #{category.chop}: #{data[category.chop].to_s}" unless $quiet
      cat_page = get_page(agent, creds, url + CONFIG, {display: category, tech_hardware: 'custom_custom'})
      frm = cat_page.form('frm_extensions')
    when 'trunks'
      puts "uploading #{category.chop}: #{data['trunk_name'].to_s}" unless $quiet
      cat_page = get_page(agent, creds, url + CONFIG, {display: category, tech: tech.upcase})
      frm = cat_page.form('trunkEdit')
  end
  abort 'error: form not found' unless frm

  if $debug
    frm.fields.each { |field| puts "send_sever_request: #{field.name}: #{field.value}" }
    frm.checkboxes.each { |chkbx| puts "send_sever_request: #{chkbx.name}: #{chkbx.value}" }
    frm.radiobuttons.each { |rdb| puts "send_sever_request: #{rdb.name}: #{rdb.value}" if rdb.checked }
  end
  # Fill in the form, and submit it!
  data.each { |key, val| frm[key] = val }
  frm.submit
end

def read_file_write_server(agent, creds, url, infilename, categories)
  wb = RubyXL::Parser.parse(infilename)
  puts "Reading from #{infilename}" if $debug
  result_page = nil
  # The username and password are stored in the 'Admin' sheet of the Excel file on row 2.  However, those
  # values are only used if they haven't been supplied on the command line or in the secrets file.
  ws = wb['Admin']
  if ws
    username ||= ws[1][0]
    password ||= ws[1][1]
  end

  SUPPORTED_TABS.each do |tab|
    ws = wb[tab]
    category, tech = tab.downcase.split('.')
    next unless categories.empty? || categories.include?(category)

    if ws
      rownum = 1
      while (row = ws[rownum])
        puts "Processing tab #{tab}, row #{rownum}" if $debug
        data = {}
        colnum = 0
        # The first column is assumed to have an value in it for all valid rows. If not, skip it (skips blank trailing rows)
        if row[0] && row[0].value && !row[0].value.to_s.empty?
          while (col = row[colnum])
            abort "Error: missing column heading for sheet #{tab}, cell #{RubyXL::Reference.ind2ref(rownum, colnum)}" unless ws[0][colnum]
            key = ws[0][colnum].value
            if /(?<prefix>.*)\/(?<suffix>.*)/ =~ key # This deals with subsettings of the form "big/small"
              data[prefix] = {} unless data[prefix]
              data[prefix][suffix] = col && col.value.to_s
            else
              data[key] = col && col.value.to_s # This is the usual case
            end
            colnum += 1
          end
          result_page = fill_form_and_submit(agent, creds, url, category, tech, data)
        end
        rownum += 1
      end
    end
  end

  if result_page
    puts('Reloading...') unless $quiet
    result = agent.post url + CONFIG, 'handler' => 'reload'
    if result.header['content-type'] == 'application/json'
      puts JSON.parse(result.body)['message'] unless $quiet
    else
      abort 'fpconfig.rb: reload failed'
    end
  end
end

####
# Main program.
# This program performs bulk data entry and retrieval to a FreePBX server.
# It interacts with a FreePBX server by emulating observed browser interactions rather than by using a REST API,
# because there doesn't appear to be a usable one.
####

begin
  opts = Slop.parse banner: 'Usage: fpconfig.rb [options] FreePBX-server-url' do |o|
    o.string '-s', '--secrets', 'pathname of secrets file in YAML format (default: secrets.yml)', default: 'secrets.yml'
    o.string '-u', '--username', 'username to access the FreePBX server'
    o.string '-p', '--password', 'password to access the FreePBX server'
    o.string '-e', '--email', 'email address for setup of unconfigured server'
    o.array '-c', '--categories', 'list of categories (e.g. "trunks") to process, default: all', delimiter: ','
    o.bool '-d', '--debug', 'print verbose debug messages'
    o.bool '-q', '--quiet', 'do not print progress messages'
    o.string '-o', '--output', 'read configuration from server and write Excel format to named output file'
    o.string '-w', '--write-to-server', 'read configuration from named Excel file and write it into the FreePBX server'
    o.on '--help' do
      puts o
      exit
    end
  end
rescue Slop::Error => e
  abort "fpconfig.rb: #{e.message}"
end


$debug = opts.debug?
$quiet = opts.quiet?

config = YAML.load(File.read(opts[:secrets]))
url = opts.arguments[0] || config['url']
abort 'FreePBX server URL must be given on the command line (first argument) or in the configuration file' unless url
username = opts[:username] || config['username']
password = opts[:password] || config['password']
email = opts[:email] || config['email']
creds = { username: username, password: password, email: email }
if opts[:output] || opts[:write_to_server]
  abort 'username to access FreePBX server must be given on the command line (--username) or in the configuration file' unless username
  abort 'password to access FreePBX server must be given on the command line (--password) or in the configuration file' unless password
end

if $debug
  puts "URL: #{url}"
  puts "Username: #{username}"
  puts "Password: [redacted]"
end

ph = config['proxy_host']
pp = config['proxy_port']

unless opts[:output] || opts[:write_to_server]
  abort opts.to_s
end

agent = Mechanize.new
agent.set_proxy(ph, pp) if $debug && ph && pp

if opts[:output]
  read_server_write_file(agent, creds, url, opts[:output],
                         config['field_blacklist'], config['field_order'], opts[:categories])
end

if opts[:write_to_server]
  read_file_write_server(agent, creds, url, opts[:write_to_server], opts[:categories])
end

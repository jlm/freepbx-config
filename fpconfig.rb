###
### fpconfig.rb
### Author: John Messenger
### License: Apache 2.0 License
###

# This program can read a FreePBX server configuration from the server and write it to an Excel spreadsheet.
# It can also read the spreadsheet format and write it back into the server.
# At present it can only handle Extensions.

require 'rubygems'
require 'mechanize'
require 'yaml'
require 'json'
require 'rubyXL'
require 'slop'

CONFIG = '/admin/config.php'
AJAX = '/admin/ajax.php'

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
# fill it in, submit it and then try again.
####
def get_page(agent, username, password, url, params, referer = nil)
  page = get_with_ref(agent, url, params, referer)
  if /text\/html/ =~ page.response['content-type']
    unless page.css('div#login_form').empty?
      login_form = page.forms[0]
      login_form.username = username
      login_form.password = password
      page = agent.submit(login_form)
      abort 'login failed' unless page.css('div#login_form').empty?
      page = get_with_ref(agent, url, params, referer)
    end
  end
  page
end

####
# Call RubyXL's add_cell method on a table being built.  If this is the first table row,
# write a corresponding header row entry.
####
def ws_add_data(ws, row, col, key, val)
  if row == 1
    ws.add_cell(row-1, col, key.to_s)
  end
  ws.add_cell(row, col, val.to_s)
  puts "ws_add_data: #{key.to_s}: #{val.to_s}" if $debug
end

####
# Log in to the FreePBX server and read various parameters.  Save them into an Excel spreadsheet.
####
def read_server_write_file(agent, username, password, url, outfilename, field_blacklist, field_order)
  stage = 2

  wb = RubyXL::Workbook.new
  ws_admin = wb.add_worksheet('Admin')
  ws_admin.add_cell(0, 0, 'Admin user')
  ws_admin.add_cell(0, 1, 'Password')
  ws_admin.add_cell(1, 0, username)
  ws_admin.add_cell(1, 1, password)

  ext_page = get_page(agent, username, password, url + CONFIG, {display: :extensions})
  ext_grid_result = get_page(agent, username, password, url + AJAX,
                             {module: :core, command: :getExtensionGrid, type: :all, order: :asc}, ext_page.uri.to_s)
  ext_grid = JSON.parse(ext_grid_result.body)

  ws = wb.add_worksheet('Extensions')
  row = 1
  ext_grid.each do |ext|
    data = {}
    ext_page = get_page(agent, username, password, url + CONFIG,
                        {display: :extensions, extdisplay: ext['extension']}, ext_grid_result.uri.to_s)
    ext_page.form('frm_extensions').fields.each { |ff| data[ff.name] = ff.value }
    ext_page.form('frm_extensions').checkboxes.each { |ff| data[ff.name] = ff.value }
    ext_page.form('frm_extensions').radiobuttons.each { |ff| data[ff.name] = ff.value if ff.checked }

    field_blacklist.each { |bad| data.delete(bad) }
    col = 0
    field_order.each do |efn|
      next unless data[efn]
      ws_add_data(ws, row, col, efn, data[efn])
      data.delete(efn)
      col += 1
    end
    data.each do |key, val|
      ws_add_data(ws, row, col, key, val)
      col += 1
    end
    row += 1
  end

  #wb.delete_worksheet('Sheet1')
  wb.write(outfilename)
end

####
# Retrieve the form relating to a particular spreadsheet row from the server and fill it in with the provided data.
# Submit the form.
####
def send_server_request(agent, username, password, url, category, data)
  category.downcase!
  result_page = nil
  puts "uploading #{category.chop}: #{data[category.chop].to_s}" unless $quiet
  if category == 'extensions'
    ext_page = get_page(agent, username, password, url + CONFIG,
                        {display: :extensions, tech_hardware: 'custom_custom'})
    result_page = ext_page.form('frm_extensions') do |frm|
      if $debug
        frm.fields.each { |field| puts "send_sever_request: #{field.name}: #{field.value}" }
        frm.checkboxes.each { |chkbx| puts "send_sever_request: #{chkbx.name}: #{chkbx.value}" }
        frm.radiobuttons.each { |rdb| puts "send_sever_request: #{rdb.name}: #{rdb.value}" if rdb.checked }
      end
      # Fill in the form, and submit it!
      data.each { |key, val| frm[key] = val }
    end.submit
    result_page
  end
end

def read_file_write_server(agent, username, password, url, infilename)
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

  ['Extensions'].each do |tab|
    ws = wb[tab]
    if ws
      rownum = 1
      while (row = ws[rownum])
        puts "Processing tab #{tab}, row #{rownum}" if $debug
        data = {}
        colnum = 0
        while (col = row[colnum])
          key = ws[0][colnum].value
          if /(?<prefix>.*)\/(?<suffix>.*)/ =~ key
            data[prefix] = {} unless data[prefix]
            data[prefix][suffix] = col && col.value.to_s
          else
            data[key] = col && col.value.to_s
          end
          colnum += 1
        end
        result_page = send_server_request(agent, username, password, url, ws.sheet_name, data)
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
if opts[:output]
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
  read_server_write_file(agent, username, password, url, opts[:output], config['extn_field_blacklist'], config['extn_field_order'])
end

if opts[:write_to_server]
  read_file_write_server(agent, username, password, url, opts[:write_to_server])
end

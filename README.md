# freepbx-config
Upload and download FreePBX configurations to an Excel file

After years of manually recreating my FreePBX configuration by hand each time the computer died (Raspberry Pis seem to do this all the time),
I finally decided to try to automate the configuration.  Everything has an easy-to-use REST API these days, right?  It 
couldn't be so hard.  But I was unable to find tools to do the job or a usable API.  So, I started writing this script,
which is based on the Ruby Mechanize gem, to automate the upload and download of configuration data.

For managing this kind of data I find Excel files very convenient, so I incorporated the RubyXL gem for configuration storage.

At this stage, only Extensions can be uploaded and downloaded, and I have no idea how effectively the details are stored.
However it seems to be working for me.  Over time, my plan is to incorporate at least Trunks, Inbound and Outbound Routes
and probably Ring Groups.

The design philosophy is to include as little dependency on the internal structure of FreePBX as possible, in order to reduce
need to change the tool when FreePBX is upgraded.

Dependencies
------------
Freepbx-config is written for Ruby 2.3.1 and depends on the following Ruby gems:
* Mechanize
* rubyXL
* slop

#Require the library for SOAP
require 'rubygems'
require 'yaml'
require 'soap/wsdlDriver'
require 'digest/md5'
require 'fastercsv'

#Load our configuration file
$config = YAML.load_file("config.yml")

#Build an output array containing the hashes of each record returned by the SOAP web service of Sugar CRM
def collect_output results
  output = []
  for entry in results.entry_list
    item = {}
    for name_value in entry.name_value_list
      item[name_value.name]=name_value.value
    end
    output << item
  end
  return output
end

#Build our credetials hash to be passed to the SOAP factory and converted to XML to pass to Sugar CRM
credentials = { "user_name" => $config["username"], "password" => Digest::MD5.hexdigest($config["password"]) }
begin
  #Connect to the Sugar CRM WSDL and build our methods in Ruby
  ws_proxy = SOAP::WSDLDriverFactory.new($config["wsdl_url"]).create_rpc_driver
  
  #This may be toggled on to log XML requests/responses for debugging
  #ws_proxy.wiredump_file_base = "soap"
  
  #Login to Sugar CRM
  session = ws_proxy.login(credentials, nil)
rescue => err
  puts err
  exit
end

#Check to see we got logged in properly
if session.error.number.to_i != 0
  puts session.error.description + " (" + session.error.number + ")"
  puts "Exiting"
  exit
else
  puts "Successfully logged in"
end

#Build our query for leads
module_name = "Leads"
query = "leads.lead_source_description = 'Astricon 2008'" # gets all the acounts, you can also use SQL like "accounts.name like '%company%'"
order_by = "leads.status" # in default order. you can also use SQL like "accounts.name"
offset = 0 # I guess this is like the SQL offset
select_fields = [] #could be ['name','industry']
max_results = "100" # if set to 0 or "", this doesn't return all the results, like you'd expect, set to 100 as do not expect more, and times out with too many
deleted = 0 # whether you want to retrieve deleted records, too, we don't want to

#Query the SOAP WS of Sugar CRM for the Leads that we are interested in
begin
  results = ws_proxy.get_entry_list(session['id'], module_name, query, order_by, offset, select_fields, max_results, deleted)
rescue => err
  puts err
end

#Organize the results into a nice array of hashes to be output into our reports
leads = collect_output(results)

#Create a YAML file so we may see all of the data, make sure there is nothing else folks want
File.open($config["yaml_fn"], 'w') {|f| f.write(leads.to_yaml)}

#Create a CSV file with a subset of details that may be opened in Excel
fh = File.open($config["excel_fn"], 'w')
#Write row headers for the Leads worksheet
fh.write("Account Name,First Name,Last Name,Assigned To,Country,Status,Description")
fh.write("\n")
#Write rows for the Leads CSV file
leads.each do |lead|
  row = lead["account_name"].gsub(",", "") + "," +
        lead["first_name"] + "," +
        lead["last_name"] + "," +
        lead["assigned_user_name"]  + "," +
        lead["primary_address_country"] + "," +
        lead["status"] + "," +
        lead["description"].gsub(",", "").gsub("\n", "") + "\n"
  fh.write(row)
end
fh.close
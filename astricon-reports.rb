#!/usr/bin/env ruby

#Require the library for SOAP
require 'lib/libs.rb'
require 'pp'

#Load our configuration file
$config = YAML.load_file("config.yml")

#First lets set the name of the quarter we are in
case Time.now.at_beginning_of_quarter.month
  when 1
    this_quarter = "Q1"
  when 4
    this_quarter = "Q2"
  when 7
    this_quarter = "Q3"
  when 10
    this_quarter = "Q4"
end

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
  ws_proxy.streamhandler.client.receive_timeout = 3600
  
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

#Hash to track the totals of the forecast
totals = { "new" => 0,
           "assigned" => 0,
           "converted" => 0,
           "dead" => 0,
           "partner" => 0,
           "end-user" => 0,
           "total" => 0 }

#Now lets qualify which leads we should be using in the forecast
astricon_report_table = Table(%w[AccountName Name Owner LeadState Description Status])
#Create totals table
totals_report_table = Table(%w[New Assigned Converted Dead Total])
#Create types table
types_report_table = Table(%w[Partner End-User Undetermined Total])

leads.each do |lead|
  astricon_report_table << [ lead["account_name"],
                             lead["first_name"] + " " + lead["last_name"],
                             lead["assigned_user_name"],
                             lead["status"],
                             lead["description"],
                             lead["status_description"] ]
  case lead["status"]
    when "New"
      totals["new"] += 1
    when "Assigned"
      totals["assigned"] += 1
    when "Converted"
      totals["converted"] += 1
    when "Dead"
      totals["dead"] += 1
  end
  if lead["description"][0,7] == "Partner"
    totals["partner"] += 1
  end
  if lead["description"][0,8] == "End-user"
    totals["end-user"] += 1
  end
  totals["total"] += 1
end

totals_report_table << [ totals["new"],
                         totals["assigned"],
                         totals["converted"],
                         totals["dead"],
                         totals["total"] ]

types_report_table << [ totals["partner"],
                        totals["end-user"],
                        totals["total"] - (totals["partner"] + totals["end-user"]),
                        totals["total"] ]
                                                  
pdf_filename = "tmp/Astricon2008_" + Time.now.to_s.gsub(" ", "-") + ".pdf"
TableRenderer.render_pdf( :file => pdf_filename,
                          :report_title => "Astricon 2008 Lead Report",
                          :sales_quarter => this_quarter + "FY" + Time.now.year.to_s,
                          :data => [ astricon_report_table, totals_report_table, types_report_table ] )
                          
#Generate CSV file
csv_filename = "tmp/Astricon2008_" + Time.now.to_s.gsub(" ", "-") + ".csv"
File.open(csv_filename, "w") do |outfile|
  outfile.puts astricon_report_table.to_csv
end
                          
puts "Completed"
# Leaseweb servers information
# Dimofinf, Inc
# v1.0
###############################
import requests
import json
import xlsxwriter

apiurl = 'https://api.leaseweb.com/v1/bareMetals'
apikey = ''
headers = {"X-Lsw-Auth": apikey}

# Final file delivered
export_file = "servers.xlsx"

# Open XLSX file and adding sheets
workbook = xlsxwriter.Workbook(export_file)

# Formatting CELLS
cells_titles_format = workbook.add_format({'bold': True})
cells_titles_format.set_bg_color("#81BEF7")

worksheet_servers_details = workbook.add_worksheet("Servers Details")

# Declaring Awesome variables for columns
baremetal_name_col = 0
baremetal_id_col = 1
baremetal_startdate_col = 2
baremetal_contract_col = 3
baremetal_hardware_col = 4

# Set width for column
worksheet_servers_details.set_column(baremetal_name_col, baremetal_name_col, 16)
worksheet_servers_details.set_column(baremetal_id_col, baremetal_id_col, 10)
worksheet_servers_details.set_column(baremetal_startdate_col, baremetal_startdate_col, 13)
worksheet_servers_details.set_column(baremetal_contract_col, baremetal_contract_col, 16)
worksheet_servers_details.set_column(baremetal_hardware_col, baremetal_hardware_col, 40)

# Format to write ( row, column, content, format )
worksheet_servers_details.write(0, baremetal_name_col, "ServerName", cells_titles_format)
worksheet_servers_details.write(0, baremetal_id_col, "ServerID", cells_titles_format)
worksheet_servers_details.write(0, baremetal_startdate_col, "StartDate", cells_titles_format)
worksheet_servers_details.write(0, baremetal_contract_col, "Contract Length", cells_titles_format)
worksheet_servers_details.write(0, baremetal_hardware_col, "Hardware", cells_titles_format)

# Get list of servers
try:
    response = requests.get(apiurl, headers=headers)
    content = response.text
    content_json = json.loads(content)
    servers_number = len(content_json["bareMetals"])

    for count in range(servers_number):
        baremetal_json = content_json["bareMetals"][count]
        baremetal_id = baremetal_json['bareMetal']["bareMetalId"]
        baremetal_name = baremetal_json['bareMetal']["serverName"]

        print("Generating information of :" + baremetal_id)

        apiurl_server = 'https://api.leaseweb.com/v1/bareMetals/' + baremetal_id
        response_server = requests.get(apiurl_server, headers=headers)
        content_server = response_server.text
        content_json_server = json.loads(content_server)

        try:
            baremetal_startdate = content_json_server['bareMetal']['serverHostingPack']['startDate']
            baremetal_contractTerm = content_json_server['bareMetal']['serverHostingPack']['contractTerm']
            baremetal_hardware = str(content_json_server['bareMetal']['server'])

            worksheet_servers_details.write(count+1, baremetal_name_col, baremetal_name)
            worksheet_servers_details.write(count+1, baremetal_id_col, baremetal_id)

            worksheet_servers_details.write(count+1, baremetal_startdate_col, baremetal_startdate)
            worksheet_servers_details.write(count+1, baremetal_contract_col, baremetal_contractTerm)
            worksheet_servers_details.write(count+1, baremetal_hardware_col, baremetal_hardware)
        except:
            pass

except IOError:
    pass

# Close the final XLSX file
workbook.close()
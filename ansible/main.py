""" Copyright (c) 2022 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
           https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""

import meraki
import pprint as pp
import xlrd
import json

excel_file = ""
meraki_api_key = ""
org_id = ""

blacklist = []

# read the workbook, we'll just use the first sheet only
wb = xlrd.open_workbook(excel_file)
sheet = wb.sheet_by_index(0)
serial = ""
ports = ""

# find the column numbers for the serial, network, and tags
mac_list = [] 

for i in range(sheet.nrows):
    mac = sheet.cell_value(i, 0).lower()
    mac_list.append(str(mac))
    #print(str(mac).strip('sep'))


dashboard = meraki.DashboardAPI(api_key=meraki_api_key, print_console=True,output_log=True,log_path="logs")
my_orgs = dashboard.organizations.getOrganizations()
input_org_id = org_id

networks = dashboard.organizations.getOrganizationNetworks(organizationId=input_org_id)

# Data to be written
json_output = []
trunk_output = [] 

for network in networks:
    clients = dashboard.networks.getNetworkClients(network["id"],total_pages=-1)
    for client in clients:
        if client['switchport'] != None:
            print(client)
            network_device = dashboard.devices.getDevice(client['recentDeviceSerial'])
            if network_device['model'] in blacklist:
                continue
            print(network_device['model'])
            try:
                port = dashboard.switch.getDeviceSwitchPort(serial=client['recentDeviceSerial'],portId=client['switchport'])
            except:
                continue
            if port['type'] != 'trunk':
                print(port)
                if client["mac"] in mac_list:
                    ports_list = []
                    
                    ports_list.append(client["switchport"])

                    ports = {
                        "ports": port
                    }

                    try:
                        dashboard.switch.cycleDeviceSwitchPorts(serial=network_device["serial"],ports=ports_list)
                        temp = {"switch":network_device["serial"],"port":client["switchport"]}
                        json_output.append(temp)
                    except:
                        continue
            if port['type'] == 'trunk':
                temp = {"switch":network_device["serial"],"port":client["switchport"]}
                trunk_output.append(temp) 
    # Serializing json
    json_object_ports = json.dumps(json_output, indent=4)
    json_object_trunks = json.dumps(trunk_output, indent=4)
    
    # Writing to sample.json
    with open("ports.json", "w") as outfile:
        outfile.write(json_object_ports)
    
    with open("trunks.json", "w") as outfile:
        outfile.write(json_object_trunks)
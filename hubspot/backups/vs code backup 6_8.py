import requests as req
import datetime
import time
import json
import pandas as pd

# keys
hapi_key = 'xxxxxxxxxxxxxxxx'

# req url
req_url = 'https://api.hubapi.com/contacts/v1/lists/all/contacts/recent?hapikey=%s\
&count=250&&property=lead_origin&property=firstname&property=lastname&property=createdate\
&property=hs_lead_status&property=hubspot_owner_id&property=hs_searchable_calculated_phone_number\
&property=partner_type&property=lead_type' % (hapi_key)

# assign req url to variable
response = req.get(req_url).json()

# # pretty print response
# print(json.dumps(response, indent=4, sort_keys=True))

# empty list to contain dictionaries of responses, to be converted to data frame and CSV
filtered_contacts = []

for x in range(len(response['contacts'])):
    entry = {}
    timestamp = datetime.datetime.fromtimestamp(int(response['contacts'][x]['properties']['createdate']['value'])/1000).strftime('%Y-%m-%d')
    timestamp = time.mktime(datetime.datetime.strptime(timestamp, '%Y-%m-%d').timetuple())
    timestamp = datetime.datetime.fromtimestamp(timestamp)
    if (timestamp > datetime.datetime.today() - datetime.timedelta(days=22)) and (timestamp < datetime.datetime.today() - datetime.timedelta(days=13)):
        try:
            if response['contacts'][x]['properties']['lead_origin']['value'] != 'NA':
                entry['Contact ID'] = response['contacts'][x]['canonical-vid']
                entry['Create Date'] = timestamp.strftime('%Y-%m-%d')
                try:
                    entry['Lead Type'] = response['contacts'][x]['properties']['lead_type']['value']
                except:
                    entry['Lead Type'] = 'N/A'
                try:
                    entry['Partner Type'] = response['contacts'][x]['properties']['partner_type']['value']
                except:
                    entry['Partner Type'] = 'N/A'
                try:
                    entry['Phone Number'] = response['contacts'][x]['properties']['hs_searchable_calculated_phone_number']['value']
                except:
                    entry['Phone Number'] = 'N/A'
                try:
                    entry['Contact Owner ID'] = response['contacts'][x]['properties']['hubspot_owner_id']['value']
                except:
                    entry['Contact Owner ID'] = 'Unassigned'
                try:
                    entry['Lead Origin'] = response['contacts'][x]['properties']['lead_origin']['value']
                except:
                    entry['Lead Origin'] = 'N/A'
                try:
                    entry['First Name'] = response['contacts'][x]['properties']['firstname']['value']
                except:
                    entry['First Name'] = 'N/A'
                try:
                    entry['Last Name'] = response['contacts'][x]['properties']['lastname']['value']
                except:
                    entry['Last Name'] = 'N/A'
                try:
                    entry['Lead Status'] = response['contacts'][x]['properties']['hs_lead_status']['value']
                except:
                    entry['Lead Status'] = 'N/A'
                if response['contacts'][x]['identity-profiles'][0]['identities'][0]['type'] == 'EMAIL':
                    entry['Email'] = response['contacts'][x]['identity-profiles'][0]['identities'][0]['value']
                else:
                    entry['Email'] = 'N/A'
                filtered_contacts.append(entry)
        except:
            entry['Contact ID'] = response['contacts'][x]['canonical-vid']
            entry['Create Date'] = datetime.datetime.fromtimestamp(int(response['contacts'][x]['properties']['createdate']['value'])/1000).strftime('%Y-%m-%d')
            try:
                entry['Lead Type'] = response['contacts'][x]['properties']['lead_type']['value']
            except:
                entry['Lead Type'] = 'N/A'
            try:
                entry['Partner Type'] = response['contacts'][x]['properties']['partner_type']['value']
            except:
                entry['Partner Type'] = 'N/A'
            try:
                entry['Phone Number'] = response['contacts'][x]['properties']['hs_searchable_calculated_phone_number']['value']
            except:
                entry['Phone Number'] = 'N/A'
            try:
                entry['Contact Owner ID'] = response['contacts'][x]['properties']['hubspot_owner_id']['value']
            except:
                entry['Contact Owner ID'] = 'Unassigned'
            try:
                entry['Lead Origin'] = response['contacts'][x]['properties']['lead_origin']['value']
            except:
                entry['Lead Origin'] = 'N/A'
            try:
                entry['First Name'] = response['contacts'][x]['properties']['firstname']['value']
            except:
                entry['First Name'] = 'N/A'
            try:
                entry['Last Name'] = response['contacts'][x]['properties']['lastname']['value']
            except:
                entry['Last Name'] = 'N/A'
            try:
                entry['Lead Status'] = response['contacts'][x]['properties']['hs_lead_status']['value']
            except:
                entry['Lead Status'] = 'N/A'
            if response['contacts'][x]['identity-profiles'][0]['identities'][0]['type'] == 'EMAIL':
                entry['Email'] = response['contacts'][x]['identity-profiles'][0]['identities'][0]['value']
            else:
                entry['Email'] = 'N/A'
            filtered_contacts.append(entry)

# print(filtered_contacts)

# convert list to data frame
df = pd.DataFrame(filtered_contacts)

# export data frame as CSV
df.to_excel('hubcontacts.xlsx', index=False)

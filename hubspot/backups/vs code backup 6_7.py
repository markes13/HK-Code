import requests as req
import datetime
import time
import json

# keys
hapi_key = 'xxxxxxxxxxxxxxxxx'

# req url
req_url = 'https://api.hubapi.com/contacts/v1/lists/all/contacts/recent?hapikey=%s\
&count=100&&property=lead_origin&property=firstname&property=lastname' % (hapi_key)

# test url
# req_url = 'https://api.hubapi.com/contacts/v1/lists/all/contacts/recent?count=100&hapikey=%s' % (hapi_key)

# assign req url to variable
response = req.get(req_url).json()

# # pretty print response
# print(json.dumps(response, indent=4, sort_keys=True))

# empty dictionary to hold response fields for later conversion to CSV/Excel
filtered_contacts_dictionary = {}

# for loop to conditionally append dictionary based on response
for x in range(len(response['contacts'])):
    timestamp = datetime.datetime.fromtimestamp(response['contacts'][x]['identity-profiles'][0]['identities'][0]['timestamp']/1000).strftime('%Y-%m-%d')
    timestamp = time.mktime(datetime.datetime.strptime(timestamp, '%Y-%m-%d').timetuple())
    timestamp = datetime.datetime.fromtimestamp(timestamp)
    if timestamp > datetime.datetime.today() - datetime.timedelta(days=4):
        try:
            if response['contacts'][x]['properties']['lead_origin']['value'] != 'NA':
                filtered_contacts_dictionary['Contact ID'] = response['contacts'][x]['canonical-vid']
                filtered_contacts_dictionary['First Name'] = response['contacts'][x]['properties']['firstname']
                try:
                    filtered_contacts_dictionary['Last Name'] = response['contacts'][x]['properties']['lastname']
                except:
                    pass
                if response['contacts'][x]['identity-profiles'][0]['identities']['type'] == 'EMAIL':
                    filtered_contacts_dictionary['Email'] = response['contacts'][x]['identity-profiles'][0]['identities']['value']
                else:
                    filtered_contacts_dictionary['Email'] = 'N/A'
        except:
            filtered_contacts_dictionary['Contact ID'] = response['contacts'][x]['canonical-vid']
            filtered_contacts_dictionary['First Name'] = response['contacts'][x]['properties']['firstname']
            try:
                filtered_contacts_dictionary['Last Name'] = response['contacts'][x]['properties']['lastname']
            except:
                pass
            if response['contacts'][x]['identity-profiles'][0]['identities']['type'] == 'EMAIL':
                filtered_contacts_dictionary['Email'] = response['contacts'][x]['identity-profiles'][0]['identities']['value']
            else:
                filtered_contacts_dictionary['Email'] = 'N/A'


# print(json.dumps(filtered_contacts_dictionary, indent=4, sort_keys=True))
print('Total contacts: ' + str(len(filtered_contacts_dictionary)))
print(filtered_contacts_dictionary)

from zeep import Client
from zeep.cache import SqliteCache
from zeep.transports import Transport
from zeep.exceptions import Fault
from zeep.plugins import HistoryPlugin
from requests import Session
from requests.auth import HTTPBasicAuth
from urllib3 import disable_warnings
from urllib3.exceptions import InsecureRequestWarning
from lxml import etree
from collections import defaultdict
import pandas as pd

disable_warnings(InsecureRequestWarning)

## USER INPUT FOR CUCM

host = "192.168.1.7"
username = "ucmadministrator"
password = "H0m3L@b!"
print("Working...")

wsdl = 'axlsqltoolkit/schema/11.5/AXLAPI.wsdl'
location = 'https://{host}:8443/axl/'.format(host=host)
binding = "{http://www.cisco.com/AXLAPIService/}AXLAPIBinding"

session = Session()
session.verify = False
session.auth = HTTPBasicAuth(username, password)

transport = Transport(cache=SqliteCache(), session=session, timeout=20)
history = HistoryPlugin()
client = Client(wsdl=wsdl, transport=transport, plugins=[history])
service = client.create_service(binding, location)

## IMPORT FROM USER_IMPORT.xlsx

user_import_data = pd.read_excel("User_Import.xlsx")

## ADD DN

index = 0
while index < len(user_import_data.index):
    dn = user_import_data.at[index, 'DIRECTORYNUMBER']
    desc = user_import_data.at[index, 'DESCRIPTION']
    part = user_import_data.at[index, 'PARTITION']
    line = {
        'pattern': dn,
        'description': desc,
        'usage': 'Device',
        'routePartitionName': part,
        'alertingName': desc,
        'asciiAlertingName': desc
    }
    try:
        service.addLine(line)
    except:
        pass        
    index += 1

## ADD DEVICE PER DEVICE TYPE
index = 0

while index < len(user_import_data.index):

    ## ADD DEVICE FOR DESKTOP 
 
    if user_import_data.at[index, 'DEVICETYPE'] == 'DESKTOP':
        name = user_import_data.at[index, 'DEVICENAME']
        css = user_import_data.loc[index, 'DEVICECSS']
        dp = user_import_data.loc[index, 'DEVICEPOOL']
        dn = user_import_data.at[index, 'DIRECTORYNUMBER']
        part = user_import_data.at[index, 'PARTITION']
        desc = user_import_data.at[index, 'DESCRIPTION']
        phone = {
            'name': 'CSF' + name,
            'description': "For Jabber Use",
            'product': 'Cisco Unified Client Services Framework',
            'model': 'Cisco Unified Client Services Framework',
            'class': 'Phone',
            'protocol': 'SIP',
            'protocolSide': 'User',
            'callingSearchSpaceName': css,
            'devicePoolName': dp,
            'commonPhoneConfigName': 'Standard Common Phone Profile',
            'locationName': 'Hub_None',
            'useTrustedRelayPoint': 'Default',
            'builtInBridgeStatus': 'Default',
            'packetCaptureMode': 'None',
            'certificateOperation': 'No Pending Operation',
            'deviceMobilityMode': 'Default',
            'lines': {
                'line': [
                    {
                        'index': 1,
                        'dirn': {
                            'pattern': dn,
                            'routePartitionName': part
                        }
                    }
                ]
            }
        }
        try:
            service.addPhone(phone)
        except:
            pass

        ## ADD USER FOR DESKTOP      

        device = user_import_data.at[index, 'DEVICENAME']
        associated_device = {
            'device': []
        }
        associated_device['device'].append('CSF' + user_import_data.at[index, 'DEVICENAME'])
        userid = user_import_data.at[index, 'USER']
        first = user_import_data.at[index, 'FIRST']
        last = user_import_data.at[index, 'LAST']
        pw = user_import_data.at[index, 'PASSWORD']
        try:
            dn = int(user_import_data.at[index, 'DIRECTORYNUMBER'])
        except:
            pass
        end_user = {
            'userid': userid,
            'firstName': first,
            'lastName': last,
            'password': pw,
            'telephoneNumber': dn,
            'presenceGroupName': 'Standard Presence Group',
            'associatedDevices': associated_device
                }
        try:
            service.addUser(end_user)
        except:
            pass

    ## ADD DEVICE FOR TABLET

    if user_import_data.at[index, 'DEVICETYPE'] == 'TABLET':
        name = user_import_data.at[index, 'DEVICENAME']
        css = user_import_data.loc[index, 'DEVICECSS']
        dp = user_import_data.loc[index, 'DEVICEPOOL']
        dn = user_import_data.at[index, 'DIRECTORYNUMBER']
        part = user_import_data.at[index, 'PARTITION']
        desc = user_import_data.at[index, 'DESCRIPTION']
        phone = {
            'name': 'TAB' + name,
            'description': "For Jabber Use - Tablet",
            'product': 'Cisco Jabber for Tablet',
            'model': 'Cisco Jabber for Tablet',
            'class': 'Phone',
            'protocol': 'SIP',
            'protocolSide': 'User',
            'callingSearchSpaceName': css,
            'devicePoolName': dp,
            'commonPhoneConfigName': 'Standard Common Phone Profile',
            'locationName': 'Hub_None',
            'useTrustedRelayPoint': 'Default',
            'builtInBridgeStatus': 'Default',
            'packetCaptureMode': 'None',
            'certificateOperation': 'No Pending Operation',
            'deviceMobilityMode': 'Default',
            'lines': {
                'line': [
                    {
                        'index': 1,
                        'dirn': {
                            'pattern': dn,
                            'routePartitionName': part
                        }
                    }
                ]
            }
        }
        try:
            service.addPhone(phone)
        except:
            pass

        ## ADD USER FOR TABLET      

        device = user_import_data.at[index, 'DEVICENAME']
        associated_device = {
            'device': []
        }
        associated_device['device'].append('TAB' + user_import_data.at[index, 'DEVICENAME'])
        userid = user_import_data.at[index, 'USER']
        first = user_import_data.at[index, 'FIRST']
        last = user_import_data.at[index, 'LAST']
        pw = user_import_data.at[index, 'PASSWORD']
        dn = user_import_data.at[index, 'DIRECTORYNUMBER']
        end_user = {
            'userid': userid,
            'firstName': first,
            'lastName': last,
            'password': pw,
            'telephoneNumber': dn,
            'presenceGroupName': 'Standard Presence Group',
            'associatedDevices': associated_device
                }
        try:
            service.addUser(end_user)
        except:
            pass


    ## ADD DEVICE FOR IPHONE

    if user_import_data.at[index, 'DEVICETYPE'] == 'IPHONE':
        name = user_import_data.at[index, 'DEVICENAME']
        css = user_import_data.loc[index, 'DEVICECSS']
        dp = user_import_data.loc[index, 'DEVICEPOOL']
        dn = user_import_data.at[index, 'DIRECTORYNUMBER']
        part = user_import_data.at[index, 'PARTITION']
        desc = user_import_data.at[index, 'DESCRIPTION']
        phone = {
            'name': 'TCT' + name,
            'description': "For Jabber Use - iPhone",
            'product': 'Cisco Dual Mode for iPhone',
            'model': 'Cisco Dual Mode for iPhone',
            'class': 'Phone',
            'protocol': 'SIP',
            'protocolSide': 'User',
            'callingSearchSpaceName': css,
            'devicePoolName': dp,
            'commonPhoneConfigName': 'Standard Common Phone Profile',
            'locationName': 'Hub_None',
            'useTrustedRelayPoint': 'Default',
            'builtInBridgeStatus': 'Default',
            'packetCaptureMode': 'None',
            'certificateOperation': 'No Pending Operation',
            'deviceMobilityMode': 'Default',
            'lines': {
                'line': [
                    {
                        'index': 1,
                        'dirn': {
                            'pattern': dn,
                            'routePartitionName': part
                        }
                    }
                ]
            }
        }

        try:
            service.addPhone(phone)
        except:
            pass

        ## ADD USER FOR IPHONE      

        device = user_import_data.at[index, 'DEVICENAME']
        associated_device = {
            'device': []
        }
        associated_device['device'].append('TCT' + user_import_data.at[index, 'DEVICENAME'])
        userid = user_import_data.at[index, 'USER']
        first = user_import_data.at[index, 'FIRST']
        last = user_import_data.at[index, 'LAST']
        pw = user_import_data.at[index, 'PASSWORD']
        try:
            dn = int(user_import_data.at[index, 'DIRECTORYNUMBER'])
        except:
            pass
        end_user = {
            'userid': userid,
            'firstName': first,
            'lastName': last,
            'password': pw,
            'telephoneNumber': dn,
            'presenceGroupName': 'Standard Presence Group',
            'associatedDevices': associated_device
                }
        try:
            service.addUser(end_user)
        except:
            pass

    ## ADD DEVICE FOR ANDROID

    if user_import_data.at[index, 'DEVICETYPE'] == 'ANDROID':
        name = user_import_data.at[index, 'DEVICENAME']
        css = user_import_data.loc[index, 'DEVICECSS']
        dp = user_import_data.loc[index, 'DEVICEPOOL']
        dn = user_import_data.at[index, 'DIRECTORYNUMBER']
        part = user_import_data.at[index, 'PARTITION']
        desc = user_import_data.at[index, 'DESCRIPTION']
        phone = {
            'name': 'BOT' + name,
            'description': "For Jabber Use - Android",
            'product': 'Cisco Dual Mode for Android',
            'model': 'Cisco Dual Mode for Android',
            'class': 'Phone',
            'protocol': 'SIP',
            'protocolSide': 'User',
            'callingSearchSpaceName': css,
            'devicePoolName': dp,
            'commonPhoneConfigName': 'Standard Common Phone Profile',
            'locationName': 'Hub_None',
            'useTrustedRelayPoint': 'Default',
            'builtInBridgeStatus': 'Default',
            'packetCaptureMode': 'None',
            'certificateOperation': 'No Pending Operation',
            'deviceMobilityMode': 'Default',
            'lines': {
                'line': [
                    {
                        'index': 1,
                        'dirn': {
                            'pattern': dn,
                            'routePartitionName': part
                        }
                    }
                ]
            }
        }

        try:
            service.addPhone(phone)
        except:
            pass

        ## ADD USER FOR IPHONE      

        device = user_import_data.at[index, 'DEVICENAME']
        associated_device = {
            'device': []
        }
        associated_device['device'].append('BOT' + user_import_data.at[index, 'DEVICENAME'])
        userid = user_import_data.at[index, 'USER']
        first = user_import_data.at[index, 'FIRST']
        last = user_import_data.at[index, 'LAST']
        pw = user_import_data.at[index, 'PASSWORD']
        try:
            dn = int(user_import_data.at[index, 'DIRECTORYNUMBER'])
        except:
            pass
        end_user = {
            'userid': userid,
            'firstName': first,
            'lastName': last,
            'password': pw,
            'telephoneNumber': dn,
            'presenceGroupName': 'Standard Presence Group',
            'associatedDevices': associated_device
                }
        try:
            service.addUser(end_user)
        except:
            pass
    else:
        print(user_import_data.at[index, 'USER'] + " was not created. Is likely a duplicate.")
    index += 1

    ## ADD USER FOR ANDROID

## UPDATE HOME CLUSTER AND ENABLE FOR IMP

for user in user_import_data['USER']:
    try:
        service.updateUser(userid=user, homeCluster=True, imAndPresenceEnable=True)
    except:
        pass

## ADD USER GROUP
# members = []
# index = 0
# while index < len(user_import_data.index):
#     members.append(user_import_data.at[index, 'USER'])
#     index += 1
# user_roles = ['Standard CCM End Users', 'Standard CTI Enabled', 'Standard CTI Allow Control of All Devices']
# jabber_user_group = {
#     'name': 'Jabber_Users',
#     'userRoles': user_roles,
#     'members': members

# }
# service.addUserGroup(jabber_user_group)

print("Complete!")

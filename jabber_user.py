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

# index = 0
# while index < len(user_import_data.index):
#     dn = user_import_data.at[index, 'DIRECTORYNUMBER']
#     desc = user_import_data.at[index, 'DESCRIPTION']
#     part = user_import_data.at[index, 'PARTITION']
#     line = {
#         'pattern': dn,
#         'description': desc,
#         'usage': 'Device',
#         'routePartitionName': part,
#         'alertingName': desc,
#         'asciiAlertingName': desc
#     }
#     try:
#         service.addLine(line)
#     except:
#         pass        
#     index += 1

## ADD CSF DEVICE PER USER

# index = 0
# while index < len(user_import_data.index):
#     name = user_import_data.at[index, 'CSFNAME']
#     css = user_import_data.loc[index, 'DEVICECSS']
#     dp = user_import_data.loc[index, 'DEVICEPOOL']
#     dn = user_import_data.at[index, 'DIRECTORYNUMBER']
#     part = user_import_data.at[index, 'PARTITION']
#     phone = {
#         'name': 'CSF' + name,
#         'description': "For Jabber Use",
#         'product': 'Cisco Unified Client Services Framework',
#         'model': 'Cisco Unified Client Services Framework',
#         'class': 'Phone',
#         'protocol': 'SIP',
#         'protocolSide': 'User',
#         'callingSearchSpaceName': css,
#         'devicePoolName': dp,
#         'commonPhoneConfigName': 'Standard Common Phone Profile',
#         'locationName': 'Hub_None',
#         'useTrustedRelayPoint': 'Default',
#         'builtInBridgeStatus': 'Default',
#         'packetCaptureMode': 'None',
#         'certificateOperation': 'No Pending Operation',
#         'deviceMobilityMode': 'Default',
#         'lines': {
#             'line': [
#                 {
#                     'index': 1,
#                     'dirn': {
#                         'pattern': dn,
#                         'routePartitionName': part
#                     }
#                 }
#             ]
#         }
#     }
#     try:
#         service.addPhone(phone)
#     except:
#         pass        
#     index += 1
    
## CREATE USER

index = 0
while index < len(user_import_data.index):
    device = user_import_data.at[index, 'CSFNAME']
    associated_device = {
        'device': []
    }
    associated_device['device'].append('CSF' + user_import_data.at[index, 'CSFNAME'])
    userid = user_import_data.at[index, 'USER']
    first = user_import_data.at[index, 'FIRST']
    last = user_import_data.at[index, 'LAST']
    pw = user_import_data.at[index, 'PASSWORD']
    end_user = {
        'userid': userid,
        'firstName': first,
        'lastName': last,
        'password': pw,
        'presenceGroupName': 'Standard Presence Group',
        'associatedDevices': associated_device
    }
    service.addUser(end_user)   
    index += 1

## UPDATE HOME CLUSTER AND ENABLE FOR IMP

# for user in user_import_data['USER']:
#     try:
#         service.updateUser(userid=user, homeCluster=True, imAndPresenceEnable=True)
#     except:
#         pass

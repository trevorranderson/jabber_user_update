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
import time
from getpass import getpass

disable_warnings(InsecureRequestWarning)



## USER INPUT FOR CUCM

host = input("Publisher IP: ")
username = input("Username: ")
password = getpass("Password: ")
impquestion = input("Enable IMP? (Yes or No): ")
impanswer = impquestion.upper()
versionquestion = input("Version (xx.x): ")
versionanswer = versionquestion.upper()
#owneruserquestion = input("Owner User ID? (Yes or No): ")
#owneruseranswer = owneruserquestion.upper()
print("Working...")

wsdl = ''
if versionanswer == "10.0":
    wsdl = 'axlsqltoolkit/schema/10.0/AXLAPI.wsdl'
elif versionanswer == "10.5":
    wsdl = 'axlsqltoolkit/schema/10.5/AXLAPI.wsdl'
elif versionanswer == "11.0":
    wsdl = 'axlsqltoolkit/schema/11.0/AXLAPI.wsdl'
elif versionanswer == "11.5":
    wsdl = 'axlsqltoolkit/schema/11.5/AXLAPI.wsdl'
elif versionanswer == "12.0":
    wsdl = 'axlsqltoolkit/schema/12.0/AXLAPI.wsdl'
elif versionanswer == "12.5":
    wsdl = 'axlsqltoolkit/schema/12.5/AXLAPI.wsdl'

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

print(user_import_data)

## CREATE DN

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
        'asciiAlertingName': desc,
        'callForwardBusy': {
            'forwardToVoiceMail': True
        },
        'callForwardBusyInt': {
            'forwardToVoiceMail': True 
        },
        'callForwardNoAnswer': {
            'forwardToVoiceMail': True
        },
        'callForwardNoAnswerInt': {
            'forwardToVoiceMail': True
        },
        'callForwardNoCoverage': {
            'forwardToVoiceMail': True
        },
        'callForwardNoCoverageInt': {
            'forwardToVoiceMail': True
        },
        'callForwardOnFailure': {
            'forwardToVoiceMail': True
        },
        'callForwardNotRegistered': {
            'forwardToVoiceMail': True
        },
        'callForwardNotRegisteredInt': {
            'forwardToVoiceMail': True
        },

    }

    try:
        service.addLine(line)
    except:
        pass       
    index += 1
    if index > len(user_import_data.index):
        print("DN's are complete.")


## CHECK IF DEVICE EXISTS

index = 0

while index < len(user_import_data.index):

    cucm_device_name = ""
    if user_import_data.at[index, 'DEVICETYPE'] == 'DESKTOP':
        cucm_device_name = "CSF" + user_import_data.at[index, 'DEVICENAME']
    elif user_import_data.at[index, 'DEVICETYPE'] == 'TABLET':
        cucm_device_name = "TAB" + user_import_data.at[index, 'DEVICENAME']
    elif user_import_data.at[index, 'DEVICETYPE'] == 'IPHONE':
        cucm_device_name = "TCT" + user_import_data.at[index, 'DEVICENAME']
    elif user_import_data.at[index, 'DEVICETYPE'] == 'ANDROID':
        cucm_device_name = "BOT" + user_import_data.at[index, 'DEVICENAME']
    else:
        print("Please select an applicable Device Type. ALL CAPS is needed.")

    cucm_device_test = "Null"
    try:
        cucm_device_test = service.getPhone(name=cucm_device_name)
    except:
        pass

    ## IF DEVICE DOES NOT EXIST == CREATE

    if cucm_device_test == "Null" or cucm_device_test['return']['phone']['name'] != cucm_device_name:
        
        print(cucm_device_name + " was not found. Device will be created.")
        
        ## ADD DEVICE FOR DESKTOP 
 
        if user_import_data.at[index, 'DEVICETYPE'] == 'DESKTOP':
            name = user_import_data.at[index, 'DEVICENAME']
            css = user_import_data.loc[index, 'DEVICECSS']
            dp = str(user_import_data.loc[index, 'DEVICEPOOL'])
            dn = str(user_import_data.at[index, 'DIRECTORYNUMBER'])
            part = user_import_data.at[index, 'PARTITION']
            desc = str(user_import_data.at[index, 'DESCRIPTION'])
            last = str(user_import_data.at[index, 'LAST'])
            extMask = user_import_data.at[index, 'EXTPHONEMASK']
            phoneDesc = dn + '|' + dp + '|' + desc + " Desktop"
            phone = {
                'name': 'CSF' + name,
                'description': phoneDesc,
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
                            },
                            'label': last + " - " + dn,
                            'display': desc,
                            'displayAscii': desc,
                            'e164Mask': extMask
                        }
                    ]
                }
            }
            try:
                service.addPhone(phone)
            except:
                pass

        ## ADD DEVICE FOR TABLET

        elif user_import_data.at[index, 'DEVICETYPE'] == 'TABLET':
            name = user_import_data.at[index, 'DEVICENAME']
            css = user_import_data.loc[index, 'DEVICECSS']
            dp = str(user_import_data.loc[index, 'DEVICEPOOL'])
            dn = str(user_import_data.at[index, 'DIRECTORYNUMBER'])
            part = user_import_data.at[index, 'PARTITION']
            desc = str(user_import_data.at[index, 'DESCRIPTION'])
            last = str(user_import_data.at[index, 'LAST'])
            extMask = user_import_data.at[index, 'EXTPHONEMASK']
            phoneDesc = dn + '|' + dp + '|' + desc + " Tablet"
            phone = {
                'name': 'TAB' + name,
                'description': phoneDesc,
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
                            },
                            'label': last + " - " + dn,
                            'display': desc,
                            'displayAscii': desc,
                            'e164Mask': extMask
                        }
                    ]
                }
            }

            service.addPhone(phone)


        ## ADD DEVICE FOR IPHONE

        elif user_import_data.at[index, 'DEVICETYPE'] == 'IPHONE':
            name = user_import_data.at[index, 'DEVICENAME']
            css = user_import_data.loc[index, 'DEVICECSS']
            dp = str(user_import_data.loc[index, 'DEVICEPOOL'])
            dn = str(user_import_data.at[index, 'DIRECTORYNUMBER'])
            part = user_import_data.at[index, 'PARTITION']
            desc = str(user_import_data.at[index, 'DESCRIPTION'])
            last = str(user_import_data.at[index, 'LAST'])
            extMask = user_import_data.at[index, 'EXTPHONEMASK']
            phoneDesc = dn + '|' + dp + '|' + desc + " iPhone"
            phone = {
                'name': 'TCT' + name,
                'description': phoneDesc,
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
                            },
                            'label': last + " - " + dn,
                            'display': desc,
                            'displayAscii': desc,
                            'e164Mask': extMask
                        }
                    ]
                }
            }


            service.addPhone(phone)


        ## ADD DEVICE FOR ANDROID

        elif user_import_data.at[index, 'DEVICETYPE'] == 'ANDROID':
            name = user_import_data.at[index, 'DEVICENAME']
            css = user_import_data.loc[index, 'DEVICECSS']
            dp = str(user_import_data.loc[index, 'DEVICEPOOL'])
            dn = str(user_import_data.at[index, 'DIRECTORYNUMBER'])
            part = user_import_data.at[index, 'PARTITION']
            desc = str(user_import_data.at[index, 'DESCRIPTION'])
            last = str(user_import_data.at[index, 'LAST'])
            extMask = user_import_data.at[index, 'EXTPHONEMASK']
            phoneDesc = dn + '|' + dp + '|' + desc + " Android"
            phone = {
                'name': 'BOT' + name,
                'description': phoneDesc,
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
                            },
                            'label': last + " - " + dn,
                            'display': desc,
                            'displayAscii': desc,
                            'e164Mask': extMask
                        }
                    ]
                }
            }

            service.addPhone(phone)


    index += 1

    ## IF DEVICE EXISTS == SKIP


    
## CHECK IF USER EXISTS


index = 0

print("Devices Complete")

while index < len(user_import_data.index):
    cucm_user = "Null"
    
    try:
        cucm_user = service.getUser(userid=user_import_data.at[index, 'USER'])
    except:
        pass

    cucm_user_upper = cucm_user.upper()

    
 

    ## IF USER DOES NOT EXIST == CREATE
        
    if cucm_user_upper == "NULL" or cucm_user['return']['user']['userid'] != user_import_data.at[index, 'USER'] :
        print(user_import_data.at[index, 'USER'] + " could not be found. User will be created.")

        ## ADD USER FOR DESKTOP      

        if user_import_data.at[index, 'DEVICETYPE'] == 'DESKTOP':
            device = user_import_data.at[index, 'DEVICENAME']
            associated_device = {
                'device': []
            }
            associated_device['device'].append('CSF' + user_import_data.at[index, 'DEVICENAME'])
            userid = user_import_data.at[index, 'USER']
            first = user_import_data.at[index, 'FIRST']
            last = user_import_data.at[index, 'LAST']
            pw = user_import_data.at[index, 'PASSWORD']
            dn = user_import_data.at[index, 'DIRECTORYNUMBER']
            jabber_user = {
            'userGroup': ['Standard CTI Enabled', 'Standard CCM End Users', 'Standard CTI Allow Control of All Devices']}
            
            end_user = {
                'userid': userid,
                'firstName': first,
                'lastName': last,
                'password': pw,
                'telephoneNumber': dn,
                'presenceGroupName': 'Standard Presence Group',
                'associatedDevices': associated_device,
                    }
            try:
                service.addUser(end_user)
            except:
                print("User " + userid + " failed to add. Check spreadsheet.")

            if impanswer == "YES":
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = True, 
                associatedGroups=jabber_user)
            else:
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= False,
                imAndPresenceEnable = False, 
                associatedGroups=jabber_user)
        
        ## ADD USER FOR TABLET      

        elif user_import_data.at[index, 'DEVICETYPE'] == 'TABLET':
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
            jabber_user = {
            'userGroup': ['Standard CTI Enabled', 'Standard CCM End Users', 'Standard CTI Allow Control of All Devices']}
            
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
                print("User " + userid + " failed to add. Check spreadsheet.")

            if impanswer == "YES":
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = True, 
                associatedGroups=jabber_user)
            else:
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = False, 
                associatedGroups=jabber_user)
        

        ## ADD USER FOR IPHONE      

        elif user_import_data.at[index, 'DEVICETYPE'] == 'IPHONE':
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
            jabber_user = {
            'userGroup': ['Standard CTI Enabled', 'Standard CCM End Users', 'Standard CTI Allow Control of All Devices']}
            
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
                print("User " + userid + " failed to add. Check spreadsheet.")

            if impanswer == "YES":
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = True, 
                associatedGroups=jabber_user)
            else:
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = False, 
                associatedGroups=jabber_user)
        
        ## ADD USER FOR ANDROID      

        elif user_import_data.at[index, 'DEVICETYPE'] == 'ANDROID':
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
            
            jabber_user = {
            'userGroup': ['Standard CTI Enabled', 'Standard CCM End Users', 'Standard CTI Allow Control of All Devices']
        }
            end_user = {
                'userid': userid,
                'firstName': first,
                'lastName': last,
                'password': pw,
                'telephoneNumber': dn,
                'presenceGroupName': 'Standard Presence Group',
                'associatedDevices': associated_device}
        
            try:
                service.addUser(end_user)
            except:
                print("User " + userid + " failed to add. Check spreadsheet.")

            if impanswer == "YES":
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = True, 
                associatedGroups=jabber_user)
            else:
                service.updateUser(
                userid = user_import_data.at[index, 'USER'],
                homeCluster= True,
                imAndPresenceEnable = False, 
                associatedGroups=jabber_user)

        else:
            pass

    index += 1
        

        ## IF USER EXISTS == UPDATE

index = 0

print("User Creation Complete")

while index < len(user_import_data.index):

    cucm_user = "Null"
    try:
        cucm_user = service.getUser(userid=user_import_data.at[index, 'USER'])
    except:
        pass    
    
    #print(cucm_user)

    if cucm_user['return']['user']['userid'] == user_import_data.at[index, 'USER']:
        cucm_device_name = ""
        if user_import_data.at[index, 'DEVICETYPE'] == 'DESKTOP':
            cucm_device_name = "CSF" + user_import_data.at[index, 'DEVICENAME']
        elif user_import_data.at[index, 'DEVICETYPE'] == 'TABLET':
            cucm_device_name = "TAB" + user_import_data.at[index, 'DEVICENAME']
        elif user_import_data.at[index, 'DEVICETYPE'] == 'IPHONE':
            cucm_device_name = "TCT" + user_import_data.at[index, 'DEVICENAME']
        elif user_import_data.at[index, 'DEVICETYPE'] == 'ANDROID':
            cucm_device_name = "BOT" + user_import_data.at[index, 'DEVICENAME']
        
        associated_device = {
           'device': []
        }
        
        try:
            current_devices = cucm_user['return']['user']['associatedDevices']['device']
        except:
            current_devices = ""

        if current_devices != "":
            for device in current_devices:
                current_device = str(device)
                associated_device['device'].append(current_device)
 
        #Add our new device to the associated device list

        associated_device['device'].append(cucm_device_name)


        jabber_user = {
            'userGroup': ['Standard CTI Enabled', 'Standard CCM End Users', 'Standard CTI Allow Control of All Devices']
        }

        print(user_import_data.at[index, 'USER'] + " was already in CUCM. User will be updated with device and roles.")

        
        if impanswer == "YES":
            service.updateUser(
                userid = user_import_data.at[index, 'USER'], 
                associatedDevices=associated_device,
                homeCluster= True,
                imAndPresenceEnable = True,
                associatedGroups=jabber_user)
        else:
            service.updateUser(
                userid = user_import_data.at[index, 'USER'], 
                associatedDevices=associated_device,
                homeCluster= True,
                imAndPresenceEnable = False,
                associatedGroups=jabber_user)


    index += 1

print("Success!")

time.sleep(30)
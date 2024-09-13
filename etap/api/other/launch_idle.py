#***********************
#
# Copyright (c) 2021-2023, Operation Technology, Inc.
#
# THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO OPERATION TECHNOLOGY, INC. 
# ANY USE OF THIS PROGRAM IS SUBJECT TO THE PROGRAM SOFTWARE LICENSE AGREEMENT, 
# EXCEPT THAT THE USER MAY MODIFY THE PROGRAM FOR ITS OWN USE. 
# HOWEVER, THE PROGRAM MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS 
# WITHOUT THE PRIOR WRITTEN CONSENT OF OPERATION TECHNOLOGY, INC.
#
#***********************



import etap.api
import sys
from etap.api.other.etap_client import EtapClient
import datetime
import json

IP = ""
PortNumber = -1
ProjName = ""

if len(sys.argv) >= 3:
    IP = sys.argv[1]
    PortNumber = int(sys.argv[2])
    ProjName = sys.argv[3]
else:
    print("Error: insufficient number of arguments.")



baseAddress = "http://{}:{}".format(IP, PortNumber)

e = etap.api.connect(baseAddress)

version = e.application.version()
P_Dict = json.loads(version)  # Python Dict
version_value = P_Dict["Version"]

year = datetime.datetime.today().year

print("**********************************************************************")
print(f"                      ETAP Python API [v{version_value}]")
print(f" (c) {year} Operation Technology, Inc. All rights reserved.")
print("")
#print("Documentation https://etapdocs.z22.web.core.windows.net/python/python/")
print("**********************************************************************")

print("import etap.api")
print("import sys")
print("from etap.api.other.etap_client import EtapClient")

print("Connecting...")

print('baseAddress = ' + baseAddress)

print('e = etap.api.connect(baseAddress)')



# ping
print("Pinging...")
pingResult = e.application.ping()
if  pingResult== True:
    print('pingResult = e.application.ping()')
    print("The connection is connected.")
else:
    print('pingResult = e.application.ping()')
    print("The connection is not connected.")

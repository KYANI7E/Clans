from copy import copy
from tkinter.messagebox import askyesno
import dragon
import scriv
import os
import json
import sys
import copy
 
print("Start")
try:
    configPath = "\\".join(sys.argv[0].split("\\")[:-1])
    with open(configPath + '\config.json', 'r') as myfile:
        data=myfile.read()
    config = json.loads(data)
    clanTag = config['clanTag']
    tokens = config['keys']
    file = configPath + '\\' + config['file']
    tags = config['tags']
    
except:
    with open('config.json', 'r') as myfile:
        data=myfile.read()
    config = json.loads(data)
    clanTag = config['clanTag']
    tokens = config['keys']
    file = config['file']
    tags = config['tags']
    

statusCode = None
statusCodeW = None
statusCodeR = None
for t in tokens:
    drago = dragon.Dragon(t, clanTag)
    if not statusCode == 200:
        (clanData, statusCode) = drago.getClanInfo(clanTag)
    if not statusCodeW == 200:
        (warData, statusCodeW) = drago.getClanWarInfo(clanTag)
    if not statusCodeR == 200:
        (raidData, statusCodeR) = drago.getClanRaids(clanTag)


flag = False
if not statusCode == 200:
    print("Could not fetch clan data")
    print(clanData)
    flag = True
if not statusCodeW == 200:
    print("Could not fetch war data")
    print(warData)
    flag = True
if not statusCodeR == 200:
    print("Could not fetch raid data")
    print(raidData)
    flag = True

if flag:
    print("Exiting script")
    exit()

clanDataRaid = copy.deepcopy(clanData)
clanDataWar = copy.deepcopy(clanData)


war = scriv.Scriv(file, tags)
war.setUpMembers(clanDataWar)

war.setUpWarColumnHeaders(2, 3, 4, 5, 6, 7, 5)
war.setUpWar(warData)
war.updateWarSheet()
war.saveFile(file)


raid = scriv.Scriv(file, tags)
raid.setUpMembers(clanDataRaid)

raid.setUpRaidColumnHeaders(2,3,4,5,6,7,8,9,5,6,7)
raid.setUpRaids(raidData)
raid.updateRiadsSheet()
raid.saveFile(file)


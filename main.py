from copy import copy
from tkinter.messagebox import askyesno
import dragon
import scriv
import os
import json
import sys
import copy
import logging

logger = logging.getLogger(__name__)
logging.basicConfig(filename="clash.log", level=logging.INFO, format="%(asctime)s :: %(name)s :: %(message)s")
logging.info("Started")

# logging.basicConfig(filename='example.log', filemode='w', level=logging.DEBUG)

def handle_unhandled_exception(exc_type, exc_value, exc_traceback):
    """Handler for unhandled exceptions that will write to the logs"""
    if issubclass(exc_type, KeyboardInterrupt):
        # call the default excepthook saved at __excepthook__
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_unhandled_exception

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
    logging.critical("Could not fetch clan data. status Code {}".format(statusCode))
    logging.critical(clanData)
    flag = True
    
if not statusCodeW == 200:
    print("Could not fetch war data")
    print(warData)
    logging.critical("Could not fetch war data. Status Code {}".format(statusCodeW))
    logging.critical(warData)
    flag = True

if not statusCodeR == 200:
    print("Could not fetch raid data")
    print(raidData)
    logging.critical("Could not fetch raid data. Status Code {}".format(statusCodeR))
    logging.critical(raidData)
    flag = True

if flag:
    logging.critical("Exiting script")
    print("Exiting script")
    exit()

clanDataRaid = copy.deepcopy(clanData)
clanDataWar = copy.deepcopy(clanData)


war = scriv.Scriv(file, tags)
war.setUpMembers(clanDataWar)

war.setUpWarColumnHeaders(2, 3, 4, 5, 6, 7, 5)
war.setUpWar(warData)
war.updateWarSheet(config['warAttacks'], config['stars'], config['warTotal'])
war.saveFile(file)


raid = scriv.Scriv(file, tags)
raid.setUpMembers(clanDataRaid)

raid.setUpRaidColumnHeaders(1,2,3,4,5,6,7,8,9,5,6,7)
raid.setUpRaids(raidData)
raid.updateRiadsSheet(config['gold'], config['raidAttacks'], config['donations'], config['raidTotal'])
raid.saveFile(file)

logging.info("Exited")

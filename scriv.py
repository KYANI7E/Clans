import enum
from functools import update_wrapper
from http.client import NETWORK_AUTHENTICATION_REQUIRED
from timeit import repeat
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill , numbers, Font, Border, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side

class Scriv():
    def saveFile(self, fileName):
        print("Saving file...")
        for i in range(1, self.war.max_column + 1):
            self.war.column_dimensions[get_column_letter(i)].auto_size = True

        for i in range(1, self.capital.max_column + 1):
            self.capital.column_dimensions[get_column_letter(i)].auto_size = True

        self.wb.save(fileName)

    def __init__(self, fileName, tags):
        self.raidDates = []
        self.raidTime = ""
        self.warDates = []
        self.warTime = ""

        self.tags = tags

        self.clanData = {}
        self.clanMembers = {}

        self.outFlag = True
        self.notAttackedFlag = True

        self.numFormat = u'#,##0;'

        self.thin_border = Border(top=Side(style='thin', color='454545'))
        self.thickBorder = Border(top=Side(style='thick'))

        self.no_fill = PatternFill(fill_type=None)
        self.red = PatternFill(fgColor='FFCCCB',
                    fill_type='solid')
        self.red2 = PatternFill(fgColor='F5C2C1',
                    fill_type='solid')
        self.green = PatternFill(fgColor='90EE90',
                    fill_type='solid')
        self.green2 = PatternFill(fgColor='86E486',
                    fill_type='solid')
        self.yellow = PatternFill(fgColor='FFFFE0',
                    fill_type='solid')
        self.yellow2 = PatternFill(fgColor='F5F5D6',
                    fill_type='solid')
        self.gray = PatternFill(fgColor='D3D3D3',
                    fill_type='solid')
        self.gray2 = PatternFill(fgColor='C9C9C9',
                    fill_type='solid')

        try:
            self.wb = load_workbook(fileName)
        except:
            self.wb = Workbook()

        try:
            self.capital = self.wb['Raids']
        except:
            self.wb.active.title = 'Raids'
            self.capital = self.wb['Raids']

        try:
            self.war = self.wb["War"]
        except:
            self.wb.create_sheet('War')
            self.war = self.wb["War"]
            
        
    def setUpWarColumnHeaders(self, tag, mapPosition, name, attacks, stars, repeat, date):
        print("Settign up war sheet...")
        self.tagPosW = tag
        self.mapPositionW = mapPosition
        self.namePosW = name
        self.attacksPosW = attacks
        self.starsPosW = stars
        self.repeatPosW = repeat
        self.datePosW = date
        self.warSheetSetUp()
        

    def setUpRaidColumnHeaders(self, tag, trophies, position, name, attacks, stars, dono, donoR, repeat,
     date, totalGold, totalDono):
        print("Setting up raid sheet...")
        self.tagPosR = tag
        self.trophiesPosR = trophies
        self.positionPosR = position
        self.namePosR = name
        self.attacksPosR = attacks
        self.goldPosR = stars
        self.donationPosR = dono
        self.donationRecievedR = donoR
        self.repeatPosR = repeat
        self.datePosR = date
        self.totalGoldR = totalGold
        self.totalDonoR = totalDono
        self.raidSetUp()

    def raidSetUp(self):
        self.capital.cell(3, self.tagPosR).value = 'Tag'
        self.capital.cell(3, self.trophiesPosR).value = 'Trophies'
        self.capital.cell(3, self.namePosR).value = 'Name'
        self.capital.cell(3, self.attacksPosR).value = 'Attacks'
        self.capital.cell(3, self.goldPosR).value = 'Gold'
        self.capital.cell(3, self.donationPosR).value = 'D'
        self.capital.cell(3, self.donationRecievedR).value = 'DR'

    def warSheetSetUp(self):
        self.war.cell(3, self.tagPosW).value = 'Tag'
        self.war.cell(3, self.mapPositionW).value = 'Position'
        self.war.cell(3, self.namePosW).value = 'Name'
        self.war.cell(3, self.attacksPosW).value = 'Attacks'
        self.war.cell(3, self.starsPosW).value = 'Stars'

    def setUpMembers(self, clanData):
        totalDonations = 0
        if self.clanMembers == {}:
            tempMembers = clanData['memberList']
            for i, p in enumerate(tempMembers):
                self.clanMembers[p['tag']] = tempMembers[i]
                self.clanMembers[p['tag']]['attacks'] = 0
                self.clanMembers[p['tag']]['capitalResourcesLooted'] = 0
                self.clanMembers[p['tag']]['status'] = 'In'
                self.clanMembers[p['tag']]['attackNumber'] = None
                self.clanMembers[p['tag']]['stars'] = None
                self.clanMembers[p['tag']]['mapPosition'] = None
                totalDonations += tempMembers[i]['donations']

        self.totalDonations = totalDonations
    
    def setUpWar(self, warData):
        print("Settign up war info...")
        rowMax = self.war.max_row
        colMax = self.war.max_column

        warD = warData['preparationStartTime'][0:8]
        self.warTime = warD[0:4] + "-" + warD[4:6] + "-" + warD[6:8]
        newInfo = False


        if not self.warTime == str(self.war.cell(2, self.datePosW).value)[:10] and not self.war.cell(2, self.datePosW).value == None:
            self.warDates.append(str(self.war.cell(2, self.datePosW).value)[:10])
            newInfo = True


        for c in range(self.repeatPosW, colMax+1, 2):
            self.warDates.append(str(self.war.cell(2, c).value))

        for r in range(4, rowMax+1):
            tag = self.war.cell(r, self.tagPosW).value
            
            if not tag in self.clanMembers:
                self.clanMembers[tag] = {}
                if newInfo:
                    self.clanMembers[tag][self.warDates[0]] = [
                        self.war.cell(r, self.attacksPosW).value,
                        self.war.cell(r, self.starsPosW).value
                    ]
                self.clanMembers[tag]['name'] = self.war.cell(r, self.namePosW).value
                self.clanMembers[tag]['tag'] = self.war.cell(r, self.tagPosW).value
                # self.clanMembers[tag]['trophies'] = self.war.cell(r, self.mapPositionW).value
                self.clanMembers[tag]['status'] = 'Out'
                self.clanMembers[tag]['attackNumber'] = None
                self.clanMembers[tag]['stars'] = None
                self.clanMembers[tag]['mapPosition'] = None
            else:
                if newInfo:
                    self.clanMembers[tag][self.warDates[0]] = [
                        self.war.cell(r, self.attacksPosW).value,
                        self.war.cell(r, self.starsPosW).value
                    ]
                
            for c in range(self.repeatPosW, colMax+1, 2):
                date = self.war.cell(2, c).value
                self.clanMembers[tag][date] = [
                    self.war.cell(r, c).value,
                    self.war.cell(r, c+1).value
                ]

        for m in self.clanMembers:
            for d in self.warDates:
                if not d in self.clanMembers[m]:
                    self.clanMembers[m][d] = [None, None]
                    self.clanMembers[m]['attacks'] = 0
                    self.clanMembers[m]['capitalResourcesLooted'] = None
                    self.clanMembers[m]['status'] = 'In'
                    self.clanMembers[m]['attackNumber'] = None
                    self.clanMembers[m]['stars'] = None
                    self.clanMembers[m]['mapPosition'] = None

        self.warAmount = 0
        for m in warData['clan']['members']:
            self.warAmount += 1
            if 'attacks' in m:
                stars = 0
                for s in m['attacks']:
                    stars += s['stars']
                self.clanMembers[m['tag']]['attackNumber'] = len(m['attacks'])
                self.clanMembers[m['tag']]['stars'] = stars
                self.clanMembers[m['tag']]['mapPosition'] = m['mapPosition']
            else: 
                self.clanMembers[m['tag']]['attackNumber'] = 0
                self.clanMembers[m['tag']]['stars'] = 0
                self.clanMembers[m['tag']]['mapPosition'] = m['mapPosition']

    def setUpRaids(self, raidData):

        print("Setting up said info...")
        rowMax = self.capital.max_row
        colMax = self.capital.max_column

        self.totalAttacks = raidData['items'][0]['totalAttacks']
        self.disctrictsDestroyed = raidData['items'][0]['enemyDistrictsDestroyed']
        self.average = round(self.totalAttacks / self.disctrictsDestroyed, 2)

        self.totalGold = raidData['items'][0]['capitalTotalLoot']

        raidT = raidData['items'][0]['startTime'][0:8]
        self.raidTime = raidT[0:4] + "-" + raidT[4:6] + "-" + raidT[6:8]
        self.raidGolds = []
        self.averages = []
        self.medals = []
        newInfo = False

        if not self.raidTime == str(self.capital.cell(1, self.datePosR).value)[:10] and not self.capital.cell(1, self.datePosR).value == None:
            self.raidDates.append(str(self.capital.cell(1, self.datePosR).value)[:10])
            self.averages.append(self.capital.cell(2, self.datePosR).value)
            self.raidGolds.append(self.capital.cell(2, self.totalGoldR).value)
            self.medals.append(self.capital.cell(1, self.totalGoldR).value)
            self.capital.cell(1, self.totalGoldR).value = "(Medals)"
            newInfo = True

        for c in range(self.repeatPosR, colMax+1, 2):
            self.raidDates.append(str(self.capital.cell(1, c).value))
            self.averages.append(self.capital.cell(2, c).value)
            self.raidGolds.append(self.capital.cell(2, c+1).value)
            self.medals.append(self.capital.cell(1, c+1).value)

        for r in range(4, rowMax+1):
            tag = self.capital.cell(r, self.tagPosR).value

            if not tag in self.clanMembers:
                self.clanMembers[tag] = {}
                self.clanMembers[tag]['tag'] = self.capital.cell(r, self.tagPosR).value
                self.clanMembers[tag]['name'] = self.capital.cell(r, self.namePosR).value
                # self.clanMembers[tag]['trophies'] = self.capital.cell(r, self.mapPositionW).value
                self.clanMembers[tag]['status'] = 'Out'
                self.clanMembers[tag]['attacks'] = 0
                self.clanMembers[tag]['capitalResourcesLooted'] = 0
                self.clanMembers[tag]['donations'] = 0
                self.clanMembers[tag]['donationsReceived'] = 0
                self.clanMembers[tag]['trophies'] = self.capital.cell(r, self.trophiesPosR).value
                if newInfo:
                    self.clanMembers[tag][self.raidDates[0]] = [
                        self.capital.cell(r, self.attacksPosR).value,
                        self.capital.cell(r, self.goldPosR).value
                    ]
            else:
                if newInfo:
                    self.clanMembers[tag][self.raidDates[0]] = [
                        self.capital.cell(r, self.attacksPosR).value,
                        self.capital.cell(r, self.goldPosR).value
                    ]

            for c in range(self.repeatPosR, colMax+1, 2):
                date = self.capital.cell(1, c).value
                self.clanMembers[tag][date] = [
                    self.capital.cell(r, c).value,
                    self.capital.cell(r, c+1).value
                ]
        
        for m in self.clanMembers:
            for d in self.raidDates:
                if not d in self.clanMembers[m]:
                    self.clanMembers[m][d] = [None, None]
                    self.clanMembers[m]['attacks'] = 0
                    self.clanMembers[m]['capitalResourcesLooted'] = 0
                    self.clanMembers[m]['donations'] = 0
                    self.clanMembers[m]['donationsReceived'] = 0
                    self.clanMembers[m]['status'] = 'In'
                    self.clanMembers[m]['attackNumber'] = None
                    self.clanMembers[m]['stars'] = None
                    self.clanMembers[m]['mapPosition'] = None

        for m in raidData['items'][0]['members']:
            self.clanMembers[m['tag']]['attacks'] = m['attacks']
            self.clanMembers[m['tag']]['capitalResourcesLooted'] = m['capitalResourcesLooted']

    def updateRiadsSheet(self):
        print("Updating raid sheet...")
        self.updateRaidVals()

        tags = self.sortGold(self.clanMembers)
        # thin_border = Border(left=Side(style='thin'), 
        #                     right=Side(style='thin'), 
        #                     top=Side(style='thin'), 
        #                     bottom=Side(style='thin'))
        
        

        self.capital.column_dimensions['C'].width = 5
        for r,m in enumerate(tags,4):
            points = -1



            self.writeRank(r, self.capital)
            self.writeCell(self.clanMembers[m], r, self.tagPosR, 'tag', self.capital, tag=True)
            self.writeCell(self.clanMembers[m], r, self.trophiesPosR, 'trophies', self.capital)
            self.writeCell(self.clanMembers[m], r, self.namePosR, 'name', self.capital)
            points += self.writeCell(self.clanMembers[m], r, self.attacksPosR, 'attacks', self.capital, params=[5,3])
            points += self.writeCell(self.clanMembers[m], r, self.goldPosR, 'capitalResourcesLooted', self.capital, params=[8000,6000])
            self.writeCell(self.clanMembers[m], r, self.donationPosR, 'donations', self.capital, params=[300,100])
            self.writeCell(self.clanMembers[m], r, self.donationRecievedR, 'donationsReceived', self.capital)
            self.colorName(r, self.namePosR, points, [4,2], self.capital)
            for i, d in enumerate(self.raidDates):
                c = (((i+1)*2)+(self.repeatPosR-2))
                self.writeCell(self.clanMembers[m], r,c, d, self.capital, params=[5,3], dated = 0)
                self.writeCell(self.clanMembers[m], r,c+1, d, self.capital, params=[8000,6000], dated = 1)
            
            if self.clanMembers[m]['attacks'] == 0 and self.notAttackedFlag:
                self.lines(r, self.capital)
                self.notAttackedFlag = False

            if self.clanMembers[m]['status'] == 'Out' and self.outFlag:
                self.lines(r, self.capital)
                self.outFlag = False


    def writeRank(self, r, sheet):
        sheet.cell(r, self.positionPosR).value = r - 3
        sheet.cell(r, self.positionPosR).alignment = Alignment(horizontal='center')
        self.colorSet(self.gray, self.gray2, r, self.positionPosR, self.capital)
        # self.capital.cell(r, self.positionPosR).border = thin_border
        # if (r-3) % 10 == 0:
        #     self.capital.cell(r, self.positionPosR).border = self.thin_border

    def lines(self, r, sheet):

        colMax = self.capital.max_column
        for c in range(1, colMax+1):
            sheet.cell(r, c).border = self.thickBorder

    def updateRaidVals(self):
        self.capital.cell(2, self.totalGoldR).value = self.totalGold
        self.capital.cell(2, self.totalGoldR).number_format  = self.numFormat
        self.capital.cell(2, self.totalDonoR).value = self.totalDonations
        self.capital.cell(2, self.totalDonoR).number_format  = self.numFormat

        self.capital.cell(1, self.datePosR).value = self.raidTime
        self.capital.cell(2, self.datePosR).value = self.average

        for i,d in enumerate(self.raidDates):
            c = (((i+1)*2)+(self.repeatPosR-2))
            self.capital.cell(1, c).value = self.raidDates[i]
            self.capital.cell(2, c).value = self.averages[i]
            self.capital.cell(2, c+1).value = self.raidGolds[i]
            self.capital.cell(2, c+1).number_format  = self.numFormat
            self.capital.cell(1, c+1).value = self.medals[i]

            self.capital.cell(3, c).value = "Attacks"
            self.capital.cell(3, c+1).value = "Gold"

    def sortGold(self, members):
        temp = []

        for i, p in enumerate(members):
            temp.append(p)

        for i, p in enumerate(temp):
            best = i
            for j in range(i,len(temp)):
                if members[temp[best]]['capitalResourcesLooted'] < members[temp[j]]['capitalResourcesLooted']:
                    best = j
                elif members[temp[best]]['capitalResourcesLooted'] == members[temp[j]]['capitalResourcesLooted']:
                    if members[temp[best]]['donations'] < members[temp[j]]['donations']:
                        best = j
            
            tt = temp[i]
            temp[i] = temp[best]
            temp[best] = tt

        return temp


    def sortPositino(self, members):
        temp = [None] * self.warAmount
        trash = []

        for m in members:
            if not members[m]['mapPosition'] == None:
                temp[members[m]['mapPosition']-1] = m
            else:
                trash.append(m)

        return temp + trash

    def updateWarSheet(self):
        print("Updating war sheet")
        self.updateWarTimes()

        tags = self.sortPositino(self.clanMembers)

        for r,m in enumerate(tags,4):
            points = -1
            self.writeCell(self.clanMembers[m], r, self.tagPosW, 'tag', self.war, tag=True)
            self.writeCell(self.clanMembers[m], r, self.mapPositionW, 'mapPosition', self.war)
            self.writeCell(self.clanMembers[m], r, self.namePosW, 'name', self.war)
            points += self.writeCell(self.clanMembers[m], r, self.attacksPosW, 'attackNumber', self.war, params=[2,1])
            points += self.writeCell(self.clanMembers[m], r, self.starsPosW, 'stars', self.war, params=[6,4])
            self.colorName(r, self.namePosW, points, [4,2], self.war)
            for i, d in enumerate(self.warDates):
                c = (((i+1)*2)+(self.repeatPosW-2))
                self.writeCell(self.clanMembers[m], r,c, d, self.war, params=[2,1], dated = 0)
                self.writeCell(self.clanMembers[m], r,c+1, d, self.war, params=[6,4], dated = 1)

    def updateWarTimes(self):
        self.war.cell(2, self.datePosW).value = self.warTime
        for i,d in enumerate(self.warDates):
            c = (((i+1)*2)+(self.repeatPosW-2))
            self.war.cell(2, c).value = self.warDates[i]
            self.war.cell(3, c).value = "Attacks"
            self.war.cell(3, c+1).value = "Stars"

    def colorName(self, r, c, points, params, sheet):
        if points >= params[0]:
            self.colorSet(self.green, self.green2, r, c, sheet)
        elif points >= params[1]:
            self.colorSet(self.yellow, self.yellow2, r, c, sheet)
        elif not points == -1:
            self.colorSet(self.red, self.red2, r, c, sheet)

    def writeCell(self, member, r, c, val, sheet, params=None, tag=False, dated=-1):
        if member['tag'] in self.tags:
            # sheet.cell(r, c).border = thin_border
            sheet.cell(r, c).font = Font(bold=True)

        else:
            sheet.cell(r, c).font = Font(bold=False)
        sheet.cell(r, c).border = None

        # if (r-3) % 11 == 0:
        #     sheet.cell(r, c).border = self.thin_border

        if tag == True:
            sheet.cell(r, c).value = member[val]
            if member[val] in self.tags:
                sheet.cell(r, c).font = Font(bold=True)
            else:
                sheet.cell(r, c).font = Font(bold=False)
            if member['status'] == 'In':
                self.colorSet(self.gray, self.gray2, r, c, sheet)
            else:
                self.colorSet(self.red, self.red2, r, c, sheet)
        elif params == None:
            if val in member:
                sheet.cell(r, c).value = member[val]
                self.colorSet(self.gray, self.gray2, r, c, sheet)
        else:
            if not dated == -1:
                sheet.cell(r, c).value = member[val][dated]
                if member[val][dated] == None:
                    self.colorSet(self.gray, self.gray2, r, c, sheet)
                    return 
                if member[val][dated] >= params[0]:
                    self.colorSet(self.green, self.green2, r, c, sheet)
                    return 3
                elif member[val][dated] >= params[1]:
                    self.colorSet(self.yellow, self.yellow2, r, c, sheet)
                    return 2
                else :
                    self.colorSet(self.red, self.red2, r, c, sheet)
                    return 1
            elif not member[val] == None and val in member: 
                sheet.cell(r, c).value = member[val]
                if member[val] >= params[0]:
                    self.colorSet(self.green, self.green2, r, c, sheet)
                    return 3
                elif member[val] >= params[1]:
                    self.colorSet(self.yellow, self.yellow2, r, c, sheet)
                    return 2
                else :
                    self.colorSet(self.red, self.red2, r, c, sheet)
                    return 1
            else: 
                self.colorSet(self.gray, self.gray2, r, c, sheet)
                sheet.cell(r, c).value = None
                return 0

    def colorSet(self, color, color2, r, c, sheet):
        if r % 2 == 0:
            sheet.cell(r, c).fill = color
        else:
            sheet.cell(r, c).fill = color2
        sheet.cell(r, c).number_format  = self.numFormat


    

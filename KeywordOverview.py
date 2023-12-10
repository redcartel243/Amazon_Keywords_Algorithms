import xlwings as xw

"""
THIS CODE GENERATES THE KEYWORD OVERVIEW REPORT"""

wb = xw.Book(
    r'F:\Documents\work\Amazon documents\FR\US work\Data analysis\FR KEYWORD OVERVIEW_COMPUTER GENERATED_2022.07.09 - 2022.10.06.xlsx')
SPROWS = 2417
SDROWS = 226
sheet1 = wb.sheets['SP KEYWORD OVERVIEW']
sheet2 = wb.sheets['Sponsored Product Keyword Repor']
sheet3 = wb.sheets['Sponsored Display Targeting Rep']
sheet4 = wb.sheets['SD KEYWORD OVERVIEW']
sheet5 = wb.sheets['SP KW DUPLICATES']
sheet6 = wb.sheets['SD KW DUPLICATES']
# [SP VARIABLES]
SPcampaignNames = sheet2['E2:E{}'.format(SPROWS)].value
SPtargeting = sheet2['G2:G{}'.format(SPROWS)].value
SPmatchType = sheet2['H2:H{}'.format(SPROWS)].value
SPClicks = sheet2['J2:J{}'.format(SPROWS)].value
SPSpend = sheet2['M2:M{}'.format(SPROWS)].value
SPSales = sheet2['P2:P{}'.format(SPROWS)].value
SPACOS = sheet2['N2:N{}'.format(SPROWS)].value
SPimpression = sheet2['I2:I{}'.format(SPROWS)].value
SPadGroup = sheet2['F2:F{}'.format(SPROWS)].value
# [SD VARIABLES]
SDcampaignNames = sheet3['D2:D{}'.format(SDROWS)].value
SDtargeting = sheet3['H2:H{}'.format(SDROWS)].value
SDmatchType = sheet3['AH2:AH{}'.format(SDROWS)].value
SDClicks = sheet3['L2:L{}'.format(SDROWS)].value
SDSpend = sheet3['O2:O{}'.format(SDROWS)].value
SDSales = sheet3['V2:V{}'.format(SDROWS)].value
SDACOS = sheet3['R2:R{}'.format(SDROWS)].value
SDimpression = sheet3['J2:J{}'.format(SDROWS)].value
SDadGroup = sheet3['G2:G{}'.format(SDROWS)].value

# (METRICS CALCULATION: SP)
temporaryValue = []
print("PROCEDURE FOR SP METRICS CALCULATION BEGINS")
x = 0
# FINDING KEYWORD BY CAMPAIGN
for value in SPcampaignNames:
    if value is None:
        break
    if value not in temporaryValue:
        temporaryValue.append(value)

# APPLYING RESULT TO THE SHEET
j = 0
for value in temporaryValue:
    sheet1[1, j].value = value
    sheet1[2, j].value = [["KEYWORD", "MATCH TYPE", 'AD GROUP', 'CLICKS', 'SPEND', 'SALES', 'ACOS', 'IMPRESSION']]
    j = j + 8

# FIND KEYWORD CORRESPONDING TO EACH CAMPAIGNS
j = 0
for value in temporaryValue:

    i = 3
    for x in range(0, SPROWS-1):
        if value == SPcampaignNames[x]:
            sheet1[i, j].value = SPtargeting[x]
            sheet1[i, j + 1].value = SPmatchType[x]
            sheet1[i, j + 2].value = SPadGroup[x]
            sheet1[i, j + 3].value = SPClicks[x]
            sheet1[i, j + 4].value = SPSpend[x]
            sheet1[i, j + 5].value = SPSales[x]
            sheet1[i, j + 6].value = SPACOS[x]
            sheet1[i, j + 7].value = SPimpression[x]

            i = i + 1
    j = j + 8


# (METRICS CALCULATION: SD)
temporaryValue = []
print("PROCEDURE FOR SP METRICS CALCULATION BEGINS")
x = 0
# FINDING KEYWORD BY CAMPAIGN
for value in SDcampaignNames:
    if value is None:
        break
    if value not in temporaryValue:
        temporaryValue.append(value)

#APPLYING RESULT TO THE SHEET
j = 0
for value in temporaryValue:
    sheet4[1, j].value = value
    sheet4[2,j].value = [["KEYWORD","MATCH TYPE",'CLICKS','SPEND','SALES','ACOS','IMPRESSION']]
    j = j + 7

# FIND KEYWORD CORRESPONDING TO EACH CAMPAIGNS
j = 0
for value in temporaryValue:

    i = 3
    for x in range(0, SDROWS-1):
        if value == SDcampaignNames[x]:
            sheet4[i, j].value = SDtargeting[x]
            sheet4[i, j + 1].value = SDmatchType[x]
            sheet4[i, j + 2].value = SDClicks[x]
            sheet4[i, j + 3].value = SDSpend[x]
            sheet4[i, j + 4].value = SDSales[x]
            sheet4[i, j + 5].value = SDACOS[x]
            sheet4[i, j + 6].value = SDimpression[x]

            i = i + 1
    j = j + 7


# (FIND DUPLICATE: SP)
temporaryValue1 = []

print("PROCEDURE FOR SP DUPLICATE CALCULATION BEGINS")
x = 0
# FINDING KEYWORD BY CAMPAIGN
for value in SPcampaignNames:
    if value is None:
        break
    if value not in temporaryValue1:
        temporaryValue1.append(value)

# APPLYING RESULT TO THE SHEET
j = 0
for value in temporaryValue1:
    sheet5[1, j].value = value
    sheet5[2, j].value = [["KEYWORD", "MATCH TYPE", 'AD GROUP', 'CLICKS', 'SPEND', 'SALES', 'ACOS', 'IMPRESSION']]
    j = j + 8

# FIND KEYWORD CORRESPONDING TO EACH CAMPAIGNS
j = 0
for value1 in temporaryValue1:
    i = 2
    y = 1
    temporaryValue2 = []
    temporaryValue3 = []
    count = 0
    for x in range(0, SPROWS-1):
        if value1 == SPcampaignNames[x]:

            if [SPtargeting[x], SPmatchType[x]] in temporaryValue2:
                count = temporaryValue2.index([SPtargeting[x], SPmatchType[x]])
                sheet5[i + y, j].value = temporaryValue3[count][0]
                sheet5[i + y, j + 1].value = temporaryValue3[count][1]
                sheet5[i + y, j + 2].value = temporaryValue3[count][2]
                sheet5[i + y, j + 3].value = temporaryValue3[count][3]
                sheet5[i + y, j + 4].value = temporaryValue3[count][4]
                sheet5[i + y, j + 5].value = temporaryValue3[count][5]
                sheet5[i + y, j + 6].value = temporaryValue3[count][6]
                sheet5[i + y, j + 7].value = temporaryValue3[count][7]

                sheet5[i + 1 + y, j].value = SPtargeting[x]
                sheet5[i + 1 + y, j + 1].value = SPmatchType[x]
                sheet5[i + 1 + y, j + 2].value = SPadGroup[x]
                sheet5[i + 1 + y, j + 3].value = SPClicks[x]
                sheet5[i + 1 + y, j + 4].value = SPSpend[x]
                sheet5[i + 1 + y, j + 5].value = SPSales[x]
                sheet5[i + 1 + y, j + 6].value = SPACOS[x]
                sheet5[i + 1 + y, j + 7].value = SPimpression[x]
                y = y + 1
                i = i + 1
            temporaryValue2.append([SPtargeting[x], SPmatchType[x]])
            temporaryValue3.append(
                [SPtargeting[x], SPmatchType[x], SPadGroup[x], SPClicks[x], SPSpend[x], SPSales[x], SPACOS[x],
                 SPimpression[x]])

    j = j + 8

# (FIND DUPLICATE: SD)
temporaryValue1 = []

print("PROCEDURE FOR SD DUPLICATE CALCULATION BEGINS")
x = 0
# FINDING KEYWORD BY CAMPAIGN
for value in SDcampaignNames:
    if value is None:
        break
    if value not in temporaryValue1:
        temporaryValue1.append(value)

# APPLYING RESULT TO THE SHEET
j = 0
for value in temporaryValue1:
    sheet6[1, j].value = value
    sheet6[2, j].value = [["KEYWORD", "MATCH TYPE", 'AD GROUP', 'CLICKS', 'SPEND', 'SALES', 'ACOS', 'IMPRESSION']]
    j = j + 8

# FIND KEYWORD CORRESPONDING TO EACH CAMPAIGNS
j = 0
for value1 in temporaryValue1:
    i = 2
    y = 1
    temporaryValue2 = []
    temporaryValue3 = []
    count = 0
    for x in range(0, SDROWS-1):
        if value1 == SDcampaignNames[x]:

            if [SDtargeting[x], SDmatchType[x]] in temporaryValue2:
                count = temporaryValue2.index([SDtargeting[x], SDmatchType[x]])
                sheet6[i + y, j].value = temporaryValue3[count][0]
                sheet6[i + y, j + 1].value = temporaryValue3[count][1]
                sheet6[i + y, j + 2].value = temporaryValue3[count][2]
                sheet6[i + y, j + 3].value = temporaryValue3[count][3]
                sheet6[i + y, j + 4].value = temporaryValue3[count][4]
                sheet6[i + y, j + 5].value = temporaryValue3[count][5]
                sheet6[i + y, j + 6].value = temporaryValue3[count][6]
                sheet6[i + y, j + 7].value = temporaryValue3[count][7]

                sheet6[i + 1 + y, j].value = SDtargeting[x]
                sheet6[i + 1 + y, j + 1].value = SDmatchType[x]
                sheet6[i + 1 + y, j + 2].value = SDadGroup[x]
                sheet6[i + 1 + y, j + 3].value = SDClicks[x]
                sheet6[i + 1 + y, j + 4].value = SDSpend[x]
                sheet6[i + 1 + y, j + 5].value = SDSales[x]
                sheet6[i + 1 + y, j + 6].value = SDACOS[x]
                sheet6[i + 1 + y, j + 7].value = SDimpression[x]
                y = y + 1
                i = i + 1
            temporaryValue2.append([SDtargeting[x], SDmatchType[x]])
            temporaryValue3.append(
                [SDtargeting[x], SDmatchType[x], SDadGroup[x], SDClicks[x], SDSpend[x], SDSales[x], SDACOS[x],
                 SDimpression[x]])

    j = j + 8

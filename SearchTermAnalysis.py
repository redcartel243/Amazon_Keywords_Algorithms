import xlwings as xw
import datetime as dt

"""
THIS CODE GENERATES THE SEARCH TERM REPORT"""

wb = xw.Book(
    r'C:\Users\Red\Documents\work\Amazon documents\FR\US work\Data analysis\Python Projects\General Search Term '
    r'Analysis\FR\FR General Search Term Analysis '
    r'Template (Macro version).xlsm')
# ---------------------
# GLOBAL SEARCH TERM ANALYSIS VARIABLES
sheet1 = wb.sheets['Sponsored Product Search Term R']
sheet2 = wb.sheets['Search Term Exact']
sheet3 = wb.sheets['Search Term Phrase']
sheet4 = wb.sheets['Search Term Broad']
sheet5 = wb.sheets['Search Term per campaign']
sheet6 = wb.sheets['Search Term Frequency and metri']
# FIND NUMBER OF ROWS
Rows = sheet1['A:A'].value
Global_search_term_Valid_Rows_Count = 0
for value in Rows:
    if value is not None:
        Global_search_term_Valid_Rows_Count = Global_search_term_Valid_Rows_Count + 1
print("{0} Rows GLOBAL SEARCH TERM ANALYSIS ".format(Global_search_term_Valid_Rows_Count))
# DEFINING RANGES (ASSIGN BY RANGE FOR FASTER LOOP)
definedTargeting = sheet1['F2:F{}'.format(Global_search_term_Valid_Rows_Count)].value
matchTypes = sheet1['G2:G{}'.format(Global_search_term_Valid_Rows_Count)].value
customerSearchTerms = sheet1['H2:H{}'.format(Global_search_term_Valid_Rows_Count)].value
campaignNames = sheet1['D2:D{}'.format(Global_search_term_Valid_Rows_Count)].value
searchTermsImpression = sheet1['I2:I{}'.format(Global_search_term_Valid_Rows_Count)].value
searchTermsClicks = sheet1['J2:J{}'.format(Global_search_term_Valid_Rows_Count)].value
# STORE TEMPORARY VALUES
Global_search_term_temporaryValue = []

# ---------------------
# ASIN SEARCH TERM ANALYSIS VARIABLES
sheet7 = wb.sheets['Sponsored Product Advertised Pr']
sheet8 = wb.sheets['ASIN->Campaign']
sheet9 = wb.sheets['Asin ST Analysis']
# FIND NUMBER OF ROWS
Rows = sheet7['A:A'].value
Asin_search_term_Valid_Rows_Count = 0
for value in Rows:
    if value is not None:
        Asin_search_term_Valid_Rows_Count = Asin_search_term_Valid_Rows_Count + 1
print("{0} for ASIN SEARCH TERM ANALYSIS".format(Asin_search_term_Valid_Rows_Count))
# DEFINING RANGES (ASSIGN BY RANGE FOR FASTER LOOP)
advertisedAsin = sheet7['G2:G{}'.format(Asin_search_term_Valid_Rows_Count)].value
asinsCampaigns = sheet7['D2:D{}'.format(Asin_search_term_Valid_Rows_Count)].value
# STORE TEMPORARY VALUES
# TEMPORARY VALUE FOR ASINS
Asin_search_term_temporaryValue1 = []
# TEMPORARY VALUE FOR CAMPAIGNS
Asin_search_term_temporaryValue2 = []
# TEMPORARY VALUE AVOIDING DUPLICATE CAMPAIGNS PER ASIN
Asin_search_term_temporaryValue3 = []
# ---------------------
# GLOBAL SEARCH TERM ANALYSIS PROCESS
# (EXACT)

# TO CLEAN UP THE SHEET
sheet2.clear_contents()
print("PROCEDURE FOR EXACT TYPE BEGINS")
sheet2['A1'].value = [["TARGETING"], ["EXACT"]]
searchStart = 3
x = 0
Global_search_term_temporaryValue.clear()
# FINDING Targeting BY MATCH TYPE
for value in definedTargeting:
    if value is None:
        break
    if matchTypes[x] == "EXACT" and value not in Global_search_term_temporaryValue:
        Global_search_term_temporaryValue.append(value)
    x = x + 1
# APPLY THE ABOVE RESULT TO THE SHEET
for value in Global_search_term_temporaryValue:
    sheet2['A{0}'.format(searchStart)].value = value
    searchStart = searchStart + 1
# FIND CUSTOMER SEARCH TERMS CORRESPONDING TO EACH TARGETING
searchStart = 2
for value in Global_search_term_temporaryValue:
    j = 1
    for x in range(1, Global_search_term_Valid_Rows_Count - 1):
        if definedTargeting[x] == value and matchTypes[x] == "EXACT":
            sheet2[searchStart, j].value = customerSearchTerms[x]
            j = j + 1
    searchStart = searchStart + 1

# (PHRASE)
# TO CLEAN UP THE SHEET

sheet3.clear_contents()
print("PROCEDURE FOR PHRASE TYPE BEGINS")
sheet3['A1'].value = [["TARGETING"], ["PHRASE"]]
searchStart = 3
x = 0
Global_search_term_temporaryValue.clear()
# FINDING Targeting BY MATCH TYPE
for value in definedTargeting:
    if value is None:
        break
    if matchTypes[x] == "PHRASE" and value not in Global_search_term_temporaryValue:
        Global_search_term_temporaryValue.append(value)
    x = x + 1
# APPLY THE ABOVE RESULT TO THE SHEET
for value in Global_search_term_temporaryValue:
    sheet3['A{0}'.format(searchStart)].value = value
    searchStart = searchStart + 1
# FIND CUSTOMER SEARCH TERMS CORRESPONDING TO EACH TARGETING
searchStart = 2
for value in Global_search_term_temporaryValue:
    j = 1
    for x in range(1, Global_search_term_Valid_Rows_Count - 1):
        if definedTargeting[x] == value and matchTypes[x] == "PHRASE":
            sheet3[searchStart, j].value = customerSearchTerms[x]
            j = j + 1
    searchStart = searchStart + 1

# (BROAD)
# TO CLEAN UP THE SHEET
sheet4.clear_contents()
print("PROCEDURE FOR BROAD TYPE BEGINS")
sheet4['A1'].value = [["TARGETING"], ["BROAD"]]
searchStart = 3
x = 0
Global_search_term_temporaryValue.clear()
# FINDING Targeting BY MATCH TYPE
for value in definedTargeting:
    if value is None:
        break
    if matchTypes[x] == "EXACT" and value not in Global_search_term_temporaryValue:
        Global_search_term_temporaryValue.append(value)
    x = x + 1
# APPLY THE ABOVE RESULT TO THE SHEET
for value in Global_search_term_temporaryValue:
    sheet4['A{0}'.format(searchStart)].value = value
    searchStart = searchStart + 1
# FIND CUSTOMER SEARCH TERMS CORRESPONDING TO EACH TARGETING
searchStart = 2
for value in Global_search_term_temporaryValue:
    j = 1
    for x in range(1, Global_search_term_Valid_Rows_Count - 1):
        if definedTargeting[x] == value and matchTypes[x] == "BROAD":
            sheet4[searchStart, j].value = customerSearchTerms[x]
            j = j + 1
    searchStart = searchStart + 1

# (CAMPAIGNS)
# TO CLEAN UP THE SHEET AND INITIALIZE VARIABLES
sheet5.clear_contents()
print("PROCEDURE FOR SEARCH TERMS PER CAMPAIGN BEGINS")
sheet5['A1'].value = [["CAMPAIGN NAME"], [""]]
searchStart = 3
x = 0
Global_search_term_temporaryValue.clear()
# FIND CAMPAIGNS
for value in campaignNames:
    if value is None:
        break
    if value not in Global_search_term_temporaryValue:
        Global_search_term_temporaryValue.append(value)

    x = x + 1

# APPLY THE ABOVE RESULT TO THE SHEET
for value in Global_search_term_temporaryValue:
    sheet5['A{0}'.format(searchStart)].value = value
    searchStart = searchStart + 1
# FIND CUSTOMER SEARCH TERMS CORRESPONDING TO EACH CAMPAIGNS
searchStart = 2
for value in Global_search_term_temporaryValue:
    j = 1
    for x in range(1, Global_search_term_Valid_Rows_Count - 1):
        if campaignNames[x] == value:
            sheet5[searchStart, j].value = customerSearchTerms[x]
            j = j + 1
    searchStart = searchStart + 1

# (SEARCH TERMS METRICS)
# TO CLEAN UP THE SHEET AND INITIALIZE VARIABLES

sheet6.clear_contents()
print("PROCEDURE FOR SEARCH TERMS METRICS")
sheet6['A1'].value = [["SEARCH TERM", "FREQUENCY", "IMPRESSION", "CLICKS", "CTR"]]
searchStart = 2
Global_search_term_temporaryValue.clear()

# FINDING Search terms
for value in customerSearchTerms:
    if value is None:
        break
    if value not in Global_search_term_temporaryValue:
        Global_search_term_temporaryValue.append(value)

# Prepare a dictionary to store metrics
search_term_metrics = {term: {"frequency": 0, "impressions": 0, "clicks": 0, "ctr": 0} for term in Global_search_term_temporaryValue}

# Calculate metrics for each search term in a single loop
for x in range(1, Global_search_term_Valid_Rows_Count - 1):
    term = customerSearchTerms[x]
    if term in search_term_metrics:
        search_term_metrics[term]["frequency"] += 1
        search_term_metrics[term]["impressions"] += int(searchTermsImpression[x])
        search_term_metrics[term]["clicks"] += int(searchTermsClicks[x])

# Calculate CTR and apply results to the sheet
for term, metrics in search_term_metrics.items():
    metrics["ctr"] = metrics["clicks"] / metrics["impressions"] if metrics["impressions"] > 0 else 0
    sheet6[f'A{searchStart}'].value = [term, metrics["frequency"], metrics["impressions"], metrics["clicks"], metrics["ctr"]]
    searchStart += 1


# ---------------------
# ASIN SEARCH TERM ANALYSIS PROCESS
# (MAPPING ASINS AND CORRESPONDING CAMPAIGNS)

# TO CLEAN UP THE SHEET
sheet8.clear_contents()
print("PROCEDURE FOR MAPPING ASINS AND CORRESPONDING CAMPAIGNS BEGINS")
sheet8['A1'].value = [["ASINS"], [""]]
searchStart = 3
Asin_search_term_temporaryValue1.clear()
# FINDING ALL ASINS AND CAMPAIGNS
for value in advertisedAsin:
    if value is None:
        break
    if value not in Asin_search_term_temporaryValue1:
        Asin_search_term_temporaryValue1.append(value)
        #print(value)
for value in asinsCampaigns:
    if value is None:
        break
    if value not in Asin_search_term_temporaryValue2:
        Asin_search_term_temporaryValue2.append(value)
        #print(value)
# APPLY THE ABOVE RESULT TO THE SHEET
for value in Asin_search_term_temporaryValue1:
    sheet8['A{0}'.format(searchStart)].value = value
    searchStart = searchStart + 1
# FIND CAMPAIGNS CORRESPONDING TO EACH ASIN
searchStart = 2
for value in Asin_search_term_temporaryValue1:
    j = 1
    Asin_search_term_temporaryValue3.clear()
    for value2 in Asin_search_term_temporaryValue2:
        for x in range(1, Asin_search_term_Valid_Rows_Count - 1):
            if advertisedAsin[x] == value and asinsCampaigns[x] == value2:
                if value2 not in Asin_search_term_temporaryValue3:
                    Asin_search_term_temporaryValue3.append(value2)
                    sheet8[searchStart, j].value = value2
                    j = j + 1
    searchStart = searchStart + 1

# (MAPPING ASINS AND CORRESPONDING CAMPAIGNS SEARCH TERMS)

# TO CLEAN UP THE SHEET
sheet9.clear_contents()
print("PROCEDURE FOR MAPPING ASINS CAMPAIGNS AND CORRESPONDING SEARCH TERMS BEGINS")
searchStart = 3
searchStart = 2
y = 0
z = 0
j = 0
for value in Asin_search_term_temporaryValue1:  # FOR EACH ASINS
    Asin_search_term_temporaryValue3.clear()
    i = j + y
    j = i
    sheet9[1, i].value = value
    sheet9[2, i].value = "Search Terms"
    print(value)
    for value2 in Asin_search_term_temporaryValue2:
        for x in range(1, Asin_search_term_Valid_Rows_Count - 1):
            if advertisedAsin[x] == value and asinsCampaigns[x] == value2:
                if value2 not in Asin_search_term_temporaryValue3:
                    j = j + 1
                    #print(value2)
                    Asin_search_term_temporaryValue3.append(value2)
                    sheet9[2, j].value = value2
                    # Find corresponding search terms
                    searchStart = 2
                    z = 3
                    for v in range(1, Global_search_term_Valid_Rows_Count - 1):
                        if value2 == campaignNames[v]:
                            sheet9[z, j].value = customerSearchTerms[v]
                            #print(j)
                            z = z + 1
    y = 1

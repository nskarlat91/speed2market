# Combined python script (OB+S2M) for the data manipulation parts I+II to generate S2M report.
# Check read me @ Github for package installation and other relevant information around the script.
# Please read https://confluence.nike.com/display/OBA/Speed+2+Market for report documentation.
# Author: Skarlatos Nikolaos

import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np

## OB PART

# Input files ( 5x Booking & Tracking + Fixed Report + ODT + StoreMaster )

bt2000 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\2000.xlsx")
bt4900 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\4900.xlsx")
bt6600 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\6600.xlsx")
bt6700 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\6700.xlsx")
odt = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\ODT.xlsx")
fixed = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\NFS - Fixed report_new (JE).xlsx")
store_master = pd.read_excel(r"C:\Users\NSkarl\Box\MPO_Replen_PRD\Store Master file.xlsx", skiprows=[0])

# Calculate Total Quantity in B&Ts

bt2000["Total Quantity"] = bt2000["Reserved Quantity"] + bt2000["AFS - Assigned Fixed Quantity"]
bt4900["Total Quantity"] = bt4900["Reserved Quantity"] + bt4900["AFS - Assigned Fixed Quantity"]
bt6600["Total Quantity"] = bt6600["Reserved Quantity"] + bt6600["AFS - Assigned Fixed Quantity"]
bt6700["Total Quantity"] = bt6700["Reserved Quantity"] + bt6700["AFS - Assigned Fixed Quantity"]

# Choose required columns only from all input files

bt2000_adj = bt2000[["SHIP TO NUMBER" , "Customer PO Number", "Division", "Sales Order Number", "Req. Delivery Date", "Total Quantity", "IDP Date", "SO rcpt Dt", "Plant"]]
bt4900_adj = bt4900[["SHIP TO NUMBER" , "Customer PO Number", "Division", "Sales Order Number", "Req. Delivery Date", "Total Quantity", "IDP Date", "SO rcpt Dt", "Plant"]]
bt6600_adj = bt6600[["SHIP TO NUMBER" , "Customer PO Number", "Division", "Sales Order Number", "Req. Delivery Date", "Total Quantity", "IDP Date", "SO rcpt Dt", "Plant"]]
bt6700_adj = bt6700[["SHIP TO NUMBER" , "Customer PO Number", "Division", "Sales Order Number", "Req. Delivery Date", "Total Quantity", "IDP Date", "SO rcpt Dt", "Plant"]]
fixed_adj = fixed[["Cust Ship To Cd", "Doc Dt", "CRD Dt", "PE", "NFS Order identification", "SO Doc Hdr Nbr", "Rsrvd + Fix Qty", "IDP Date", "Cust PO Nbr", "Status", "Div + Shpg Lctn"]]
odt_adj = odt[["Final IDP date", "Delivery Window", "Total Qty", "Customer PO Number", "Division Code", "Plant Code", "Sales Order Header Creation Date PDT", "Sales Order Header Number", "Ship To Customer Number"]]
store_master_adj = store_master[["Country", "IP", "SHIP TO", "Store name", "CO Region"]]

# Rename columns

bt2000_adj.rename(columns={"SHIP TO NUMBER" : "Cust Ship To Cd", "Customer PO Number": "Cust PO Nbr", "Division": "Div nm", "Sales Order Number": "SO Doc Hdr Nbr", "Req. Delivery Date": "CRD Dt", "SO rcpt Dt": "Doc Dt", "Plant": "Plnt Id Cd"}, inplace = True)
bt4900_adj.rename(columns={"SHIP TO NUMBER" : "Cust Ship To Cd", "Customer PO Number": "Cust PO Nbr", "Division": "Div nm", "Sales Order Number": "SO Doc Hdr Nbr", "Req. Delivery Date": "CRD Dt", "SO rcpt Dt": "Doc Dt", "Plant": "Plnt Id Cd"}, inplace = True)
bt6600_adj.rename(columns={"SHIP TO NUMBER" : "Cust Ship To Cd", "Customer PO Number": "Cust PO Nbr", "Division": "Div nm", "Sales Order Number": "SO Doc Hdr Nbr", "Req. Delivery Date": "CRD Dt", "SO rcpt Dt": "Doc Dt", "Plant": "Plnt Id Cd"}, inplace = True)
bt6700_adj.rename(columns={"SHIP TO NUMBER" : "Cust Ship To Cd", "Customer PO Number": "Cust PO Nbr", "Division": "Div nm", "Sales Order Number": "SO Doc Hdr Nbr", "Req. Delivery Date": "CRD Dt", "SO rcpt Dt": "Doc Dt", "Plant": "Plnt Id Cd"}, inplace = True)
fixed_adj.rename(columns={"PE": "Div nm", "NFS Order identification": "Identification", "Rsrvd + Fix Qty": "Total Quantity", "Div + Shpg Lctn": "Plnt Id Cd"}, inplace = True)
odt_adj.rename(columns={"Final IDP date": "IDP Date", "Total Qty": "Total Quantity", "Customer PO Number": "Cust PO Nbr", "Division Code": "Div nm", "Plant Code": "Plnt Id Cd", "Sales Order Header Creation Date PDT": "Doc Dt", "Sales Order Header Number": "SO Doc Hdr Nbr", "Ship To Customer Number": "Cust Ship To Cd"}, inplace = True)

# Derive shipping point from plant code

for i in range(0, len(fixed_adj["Plnt Id Cd"])):
    fixed_adj['Plnt Id Cd'].iloc[i] = fixed_adj["Plnt Id Cd"].iloc[i][-4:]

# Identify planned with issue

for i in range(0, len(fixed_adj["Status"])):
    if fixed_adj['Status'].iloc[i] == "Planned with issue":
        fixed_adj['Status'].iloc[i] = "Planned with issue"
    else:
        fixed_adj['Status'].iloc[i] = ""

# Merge the B&Ts

merged = pd.concat([bt2000_adj, bt4900_adj, bt6600_adj, bt6700_adj], join = "outer")

# Clean the BTs from ODT values (duplicates)

cond1 = merged['SO Doc Hdr Nbr'].isin(odt['Sales Order Header Number'])
merged.drop(merged[cond1].index, inplace = True)

# Clean the fixed from B&T data (duplicates)

cond2 = fixed_adj["SO Doc Hdr Nbr"].isin(merged['SO Doc Hdr Nbr'])
fixed_adj.drop(fixed_adj[cond2].index , inplace = True)

# alternatively bt2000[~bt2000.isin(bt4900)].dropna()

# Merge ODT,BTs & Fixed

merged2 = pd.concat([merged, fixed_adj, odt_adj], join="outer")

# Define REPLEN orders

merged2['REPLEN'] = ""

for z in range(0, len(merged2["Cust PO Nbr"])):
    merged2["REPLEN"].iloc[z] = merged2["Cust PO Nbr"].iloc[z][:9]

# Define NVS order identification

merged2.Identification = ""

merged2.Identification = np.where(merged2["Cust PO Nbr"].str.contains("HASH"), "HASH", np.where(merged2["Cust PO Nbr"].str.contains("ACTV"),
"ACTIVATION", np.where(merged2["Cust PO Nbr"].str.contains("SOCKS"), "DP SOCKS", np.where(merged2["Cust PO Nbr"].str.contains("REFIT"), "REFIT",
np.where(merged2["Cust PO Nbr"].str.contains("open"), "STORE OPENING", np.where(merged2["Cust PO Nbr"].str.contains("STAFF"), "STAFF DRESS",
np.where(merged2["Cust PO Nbr"].str.contains("GLV"), "GLOBAL VISIT", np.where(merged2["Cust PO Nbr"].str.contains("BACKW"), "BACK WALL",
np.where(merged2["Cust PO Nbr"].str.contains("INIT"), "INITIAVE", np.where(merged2["REPLEN"].str.contains("R"), "REPLEN", "OTHERS"))))))))))

# Merge file with store master to get store/country/ship-to/ip/region information

final = pd.merge(merged2, store_master_adj, how="inner" , left_on="Cust Ship To Cd" , right_on = "SHIP TO")

# Get PE as str from DIVnm

final.astype({"Div nm": str})

final["Div nm"] = np.where(final["Div nm"] == 10, "APP", np.where(final["Div nm"] == 20 , "FTW", np.where(final["Div nm"] == 30 , "EQ", np.where(final["Div nm"] == "FTW" , "FTW",
np.where(final["Div nm"] == "APP" , "APP", "EQ")))))

# Filter dataframe

OB=final[final["Total Quantity"] > 0]

OB = OB[(OB["Country"] != "Turkey") & (OB["Country"] != "RUSSIA") & (OB["Country"] != "ISRAEL") & (OB["Status"] != "Planned with issue") & (OB["CO Region"] != "NSO")]
OB = OB[["CO Region","IP","Store name","Cust Ship To Cd","Cust PO Nbr","Div nm","Plnt Id Cd","SO Doc Hdr Nbr","CRD Dt","Doc Dt","IDP Date","Total Quantity","Identification","Status","Delivery Window"]]

## S2M PART

#STORE MASTER

#FPDD calculation

today = pd.Timestamp.today()

store_master["FPDD"] = ""

for i in range(0, len(store_master.iloc[:,12])):
    bd = pd.tseries.offsets.BusinessDay(n=int(store_master["PT + TT"].iloc[i]))
    store_master["FPDD"].iloc[i] = today + bd

store_master["FPDD"] = pd.to_datetime(store_master["FPDD"])
store_master["FPDD"] = store_master["FPDD"].dt.date

store_master["cap"] = ""

delivery_days = ["MON",	"TUE","WED","THU","FRI","SAT","SUN"]
store_master["number_of_deliveries"] = store_master[delivery_days].sum(axis=1)

#daily capacity calculation

for i in range(0, len(store_master["number_of_deliveries"])):
    if store_master["number_of_deliveries"].iloc[i] <= 5:
        store_master["cap"].iloc[i] = store_master["MAX CARTONS (1st window / overall)"].iloc[i]
    else:
        if store_master["MAX CARTONS (1st window / overall)"].iloc[i] == store_master["MAX CARTONS (2nd Window)"].iloc[i]:
            store_master["cap"].iloc[i] = store_master["MAX CARTONS (1st window / overall)"].iloc[i] + (store_master["number_of_deliveries"].iloc[i] - 5) \
                                          * (store_master["MAX CARTONS (1st window / overall)"].iloc[i] / 5)
        elif (store_master["MAX CARTONS (2nd Window)"].iloc[i] == 0) and (store_master["number_of_deliveries"].iloc[i] == 6):
            store_master["cap"].iloc[i] = store_master["MAX CARTONS (1st window / overall)"].iloc[i] + (store_master["MAX CARTONS (1st window / overall)"].iloc[i] / 5)
        else:
            if store_master["MAX CARTONS (2nd Window)"].iloc[i] == 0:
                store_master["cap"].iloc[i] = store_master["MAX CARTONS (1st window / overall)"].iloc[i]
            else:
                store_master["cap"].iloc[i] = store_master["MAX CARTONS (1st window / overall)"].iloc[i] + (
                 store_master["number_of_deliveries"].iloc[i] - 5) \
                    * (store_master["MAX CARTONS (2nd Window)"].iloc[i] / 5)

#OB

ob = OB

#Units per carton calculation
APP = 20
FTW = 6
EQP = 26

#Calculate cartons per line
ob["Cartons"] = ""

for i in range(0, len(ob["Div nm"])):
    if ob["Div nm"].iloc[i] == "EQ":
        ob["Cartons"].iloc[i] = ob["Total Quantity"].iloc[i] / EQP
    elif ob["Div nm"].iloc[i] == "FTW":
        ob["Cartons"].iloc[i] = ob["Total Quantity"].iloc[i] / FTW
    else:
        ob["Cartons"].iloc[i] = ob["Total Quantity"].iloc[i] / APP

#Pivot planned / unplanned

ob = pd.merge(ob, store_master[["SHIP TO", "FPDD"]], left_on="Cust Ship To Cd" , right_on="SHIP TO")

ob["IDP Date"] = pd.to_datetime(ob["IDP Date"])

u = ob[ob["IDP Date"].isnull()]
p = ob[(ob["IDP Date"] >= ob["FPDD"]) & (~ob["Identification"].isin(["ACTIVATION", "INITIATIVE"]))]

pivot_unplanned = pd.pivot_table(u, values=["Total Quantity"], index=["Cust Ship To Cd"], columns=["Div nm"], aggfunc=np.sum).add_suffix("_unplanned")
pivot_unplanned.columns.name = None
pivot_unplanned.columns = pivot_unplanned.columns.droplevel(0)

pivot_planned = pd.pivot_table(p, values=["Total Quantity"], index=["Cust Ship To Cd"], columns=["Div nm"], aggfunc=np.sum).add_suffix("_planned")
pivot_planned.columns.name = None
pivot_planned.columns = pivot_planned.columns.droplevel(0)

table = pd.concat([pivot_unplanned, pivot_planned], axis=1, join="outer")
overview_noNA = table.fillna(0)

#MERGING OB AND STORE MASTER / calculate worth ob planned/unplanned OB

overview = pd.merge(store_master[["SHIP TO", "FPDD", "cap", "IP", "EMEA Region", "Store name", "League","Subleague", "Country"]],overview_noNA, left_on="SHIP TO", right_on="Cust Ship To Cd", how="right", suffixes=(False, False))

overview['worth_of_pipeline'] = ''
overview['worth_of_scheduled'] = ''

for i in range(0, len(overview["worth_of_pipeline"])):
    overview['worth_of_pipeline'].iloc[i] = (overview["APP_unplanned"].iloc[i] / APP + overview["FTW_unplanned"].iloc[i] / FTW + overview["EQ_unplanned"].iloc[i] / EQP)\
                                            / overview["cap"].iloc[i]

for j in range(0, len(overview["worth_of_pipeline"])):
    overview['worth_of_scheduled'].iloc[j] = (overview["APP_planned"].iloc[j] / APP + overview["EQ_planned"].iloc[j] / EQP + overview["FTW_planned"].iloc[j] / FTW) \
                                             / overview["cap"].iloc[j]

overview['s2m'] = overview["worth_of_pipeline"] + overview["worth_of_scheduled"]

#SOH

sohreadapp = pd.read_excel(r'\\hilversm-nss-01\shareddata05\Retail.EHQ\MERCHAND\FACALLOCATION\Allocation analysis\SOH expectation report\Weekly SOH expectation - dist. email.xlsx',
                    sheet_name="APPAREL", skiprows=[0])
sohreadftw = pd.read_excel(r'\\hilversm-nss-01\shareddata05\Retail.EHQ\MERCHAND\FACALLOCATION\Allocation analysis\SOH expectation report\Weekly SOH expectation - dist. email.xlsx',
                    sheet_name="FOOTWEAR", skiprows=[0])
sohreadeq = pd.read_excel(r'\\hilversm-nss-01\shareddata05\Retail.EHQ\MERCHAND\FACALLOCATION\Allocation analysis\SOH expectation report\Weekly SOH expectation - dist. email.xlsx',
                    sheet_name="EQUIPMENT", skiprows=[0])

soh_app = sohreadapp.iloc[:,np.r_[0,57,58]]
sohapp = soh_app.rename(columns = {soh_app.columns[1]: "WOC APP", soh_app.columns[2]: "SOH APP"})

soh_ftw = sohreadftw.iloc[:,np.r_[0,58,59]]
sohftw = soh_ftw.rename(columns = {soh_ftw.columns[1]: "WOC FTW", soh_app.columns[2]: "SOH FTW"})

soh_eq = sohreadeq.iloc[:,np.r_[0,57,58]]
soheq = soh_eq.rename(columns = {soh_eq.columns[1]: "WOC EQ", soh_app.columns[2]: "SOH EQ"})

sohappftwmerged = pd.merge(sohapp , sohftw, on="#")
soh = pd.merge(sohappftwmerged, soheq, on="#")

#Merge & clean file

overviewsoh = pd.merge(overview, soh, left_on="IP", right_on="#", how="inner")
overviewdrop = overviewsoh.drop(columns="#")
overviewfinal = overviewdrop.round(2)

#PE Calculations

overviewfinal["Total"] = overviewfinal["APP_unplanned"] + overviewfinal["FTW_unplanned"] + overviewfinal["EQ_unplanned"] + overviewfinal["APP_planned"]\
                         +overviewfinal["FTW_planned"] + overviewfinal["EQ_planned"]

overviewfinal["APP PE%"] = (overviewfinal["APP_unplanned"] + overviewfinal["APP_planned"]) / overviewfinal["Total"]
overviewfinal["FTW PE%"] = (overviewfinal["FTW_unplanned"] + overviewfinal["FTW_planned"]) / overviewfinal["Total"]
overviewfinal["EQ PE%"] = (overviewfinal["EQ_unplanned"] + overviewfinal["EQ_planned"]) / overviewfinal["Total"]

overviewfinal["Days of receiving APP"] = overviewfinal["APP PE%"] * overviewfinal["s2m"]
overviewfinal["Days of receiving FTW"] = overviewfinal["FTW PE%"] * overviewfinal["s2m"]
overviewfinal["Days of receiving EQ"] = overviewfinal["EQ PE%"] * overviewfinal["s2m"]

#Ratio tables

overviewfinal['Ratio low woc APP'] = ''
overviewfinal['Ratio low woc FTW'] = ''
overviewfinal['Ratio high s2m APP'] = ''
overviewfinal['Ratio high s2m FTW'] = ''

for i in range(0, len(overviewfinal["WOC APP"])):
    if overviewfinal["WOC APP"].iloc[i] <= 5 and overviewfinal["Days of receiving APP"].iloc[i] <= 3 and overviewfinal["SOH APP"].iloc[i] <= 1:
        overviewfinal['Ratio low woc APP'].iloc[i] = overviewfinal["WOC APP"].iloc[i] * overviewfinal["Days of receiving APP"].iloc[i]
    else:
        overviewfinal['Ratio low woc APP'].iloc[i] = ""

for i in range(0, len(overviewfinal["WOC FTW"])):
    if overviewfinal["WOC FTW"].iloc[i] <= 7 and overviewfinal["Days of receiving FTW"].iloc[i] <= 3 and overviewfinal["SOH FTW"].iloc[i] <= 1:
        overviewfinal['Ratio low woc FTW'].iloc[i] = overviewfinal["WOC FTW"].iloc[i] * overviewfinal["Days of receiving FTW"].iloc[i]
    else:
        overviewfinal['Ratio low woc FTW'].iloc[i] = ""

for j in range(0, len(overviewfinal["WOC APP"])):
    if overviewfinal["WOC APP"].iloc[j] <= 5 and overviewfinal["Days of receiving APP"].iloc[j] >= 6:
        overviewfinal['Ratio high s2m APP'].iloc[j] = overviewfinal["WOC APP"].iloc[j] * overviewfinal["Days of receiving APP"].iloc[j]
    else:
        overviewfinal['Ratio high s2m APP'].iloc[j] = ""

for j in range(0, len(overviewfinal["WOC FTW"])):
    if overviewfinal["WOC FTW"].iloc[j] <= 7 and overviewfinal["Days of receiving FTW"].iloc[j] >= 6:
        overviewfinal['Ratio high s2m FTW'].iloc[j] = overviewfinal["WOC FTW"].iloc[j] * overviewfinal["Days of receiving FTW"].iloc[j]
    else:
        overviewfinal['Ratio high s2m FTW'].iloc[j] = ""

clearance = [509,541,557,562,588,642,644,667,677,687,690,800,896,927,934,936,945,2003,2006,2007,2008,2009,2042]

overviewfinal["Clearance?"] = ""

for x in range(0,len(overviewfinal["IP"])):
    if overviewfinal["IP"].iloc[x] in clearance:
        overviewfinal["Clearance?"].iloc[x] = "Y"
    else:
        overviewfinal["Clearance?"].iloc[x] = "N"

print(overviewfinal)

overviewfinal.to_csv(r'C:\Users\NSkarl\Box\Speed 2 Market Dashboard\speed2market2.csv')

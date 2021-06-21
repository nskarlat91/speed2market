# Python script for the data manipulation for S2M Report
# Check read me for package installation.
# author : Skarlatos Nikolaos

import pandas as pd
import numpy as np

#STORE MASTER

#fpdd calculation

store_master = pd.read_excel(r"C:\Users\NSkarl\Box\MPO_Replen_PRD\Store Master file.xlsx", skiprows=[0])

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

ob = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\OB.xlsx")

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

overviewfinal.to_csv(r'C:\Users\NSkarl\Box\Speed 2 Market Dashboard\speed2market.csv')

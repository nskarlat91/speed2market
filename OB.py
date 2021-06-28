# Python script for the data manipulation to obtain the daily NVS OrderBook
# Check read me for package installation and other relevant information.
# author : Skarlatos Nikolaos

import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import numpy as np

# Input files ( 5x Booking & Tracking + Fixed Report + ODT + Store Master )

bt2000 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\2000.xlsx")
bt4900 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\4900.xlsx")
bt6600 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\6600.xlsx")
bt6700 = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\6700.xlsx")
odt = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\ODT.xlsx")
fixed = pd.read_excel(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\Data Foundation Inputs\NFS - Fixed report_new (JE).xlsx")
store_master = pd.read_excel(r"C:\Users\NSkarl\Box\MPO_Replen_PRD\Store Master file.xlsx", skiprows=[0])

# Derive Total Quantity

bt2000["Total Quantity"] = bt2000["Reserved Quantity"] + bt2000["AFS - Assigned Fixed Quantity"]
bt4900["Total Quantity"] = bt4900["Reserved Quantity"] + bt4900["AFS - Assigned Fixed Quantity"]
bt6600["Total Quantity"] = bt6600["Reserved Quantity"] + bt6600["AFS - Assigned Fixed Quantity"]
bt6700["Total Quantity"] = bt6700["Reserved Quantity"] + bt6700["AFS - Assigned Fixed Quantity"]

# Choose required columns

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

merged2['REPLEN'] = ""

for z in range(0, len(merged2["Cust PO Nbr"])):
    merged2["REPLEN"].iloc[z] = merged2["Cust PO Nbr"].iloc[z][:9]

merged2.Identification = ""

merged2.Identification = np.where(merged2["Cust PO Nbr"].str.contains("HASH"), "HASH", np.where(merged2["Cust PO Nbr"].str.contains("ACTV"),
"ACTIVATION", np.where(merged2["Cust PO Nbr"].str.contains("SOCKS"), "DP SOCKS", np.where(merged2["Cust PO Nbr"].str.contains("REFIT"), "REFIT",
np.where(merged2["Cust PO Nbr"].str.contains("open"), "STORE OPENING", np.where(merged2["Cust PO Nbr"].str.contains("STAFF"), "STAFF DRESS",
np.where(merged2["Cust PO Nbr"].str.contains("GLV"), "GLOBAL VISIT", np.where(merged2["Cust PO Nbr"].str.contains("BACKW"), "BACK WALL",
np.where(merged2["Cust PO Nbr"].str.contains("INIT"), "INITIAVE", np.where(merged2["REPLEN"].str.contains("R"), "REPLEN", "OTHERS"))))))))))

final = pd.merge(merged2, store_master_adj, how="inner" , left_on="Cust Ship To Cd" , right_on = "SHIP TO")

final.astype({"Div nm": str})

final["Div nm"] = np.where(final["Div nm"] == 10, "APP", np.where(final["Div nm"] == 20 , "FTW", np.where(final["Div nm"] == 30 , "EQ", np.where(final["Div nm"] == "FTW" , "FTW",
np.where(final["Div nm"] == "APP" , "APP", "EQ")))))

OB=final[final["Total Quantity"] > 0]

OB = OB[OB["CO Region"] != "NSO"]
OB = OB[(OB["Country"] != "Turkey") & (OB["Country"] != "RUSSIA") & (OB["Country"] != "ISRAEL") & (OB["Status"] != "Planned with issue")]
OB = OB[["CO Region","IP","Store name","Cust Ship To Cd","Cust PO Nbr","Div nm","Plnt Id Cd","SO Doc Hdr Nbr","CRD Dt","Doc Dt","IDP Date","Total Quantity","Identification","Status","Delivery Window"]]

OB.to_csv(r"\\hilversm-nss-01\SharedData03\CustomerSvcs.Countries.EHQ\08.Strategics\04.Nike Retail\7. FACTORY STORES\27-Data Foundation\OBpython.csv", index = False)



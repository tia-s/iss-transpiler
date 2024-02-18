from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def AddFields():
wd.addCol("NARR", lambda row: "")
wd.addCol("NARR", lambda row: "")
wd.addCol("NARR", lambda row: "")
wd.addCol("NARR", lambda row: "")
wd.renameCol(columns={"BIRTHDAT": "BIRTHDAT_DATE"})
def AExt_Master():
wd.extract("Master_Funds", cols=['"UTCID"', '"TRUSTCOD"', '"ACCTNOP"'])
def BACreateDate():
wd.addCol("DATE_CREATED", lambda row: "")

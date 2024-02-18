from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def AddFields():
	wd.addCol("NARR", lambda row: "")
	wd.addCol("NARR", lambda row: "")
	wd.addCol("NARR", lambda row: "")
	wd.addCol("NARR", lambda row: "")
	wd.renameCol(columns={"BIRTHDAT": "BIRTHDAT_DATE"})
	def AExt_Master():
		extract: {'fields': ['"UTCID"', '"TRUSTCOD"', '"ACCTNOP"'], 'db_name': '"Master_Funds"', 'filter': '', 'create_virtual_database': False, 'perform_task': ''}
		def BACreateDate():
			wd.addCol("DATE_CREATED", lambda row: "")
			
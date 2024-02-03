Function Add_Account_Info_to_Transaction
If haveRecords("General Acct Master Lite.IMD")  And haveRecords("General Acct2 Master Lite.IMD")Then
	Set db = Client.OpenDatabase("General Acct Master Lite.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "FORACID"
		task.AddFieldToInc "ACID"
		task.AddFieldToInc "ACCT_NAME"
		task.AddFieldToInc "CUST_NAME"
		task.AddFieldToInc "CUST_PERM_ADDR1"
		task.AddFieldToInc "CUST_PERM_ADDR2"
		task.AddFieldToInc "SCHM_DESC"
		task.AddFieldToInc "CUSTOMER_TYPE"
		dbName = "GAM Extract.IMD"
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("General Acct Master Lite.IMD", "Add_Account_Info_to_Transaction", "Direct Extraction", "Error", "Databases empty or does not exist.")		
End If
End Function


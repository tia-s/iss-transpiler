Function VisualConnectors
If haveRecords("Cust Txn Status Change.IMD") And haveRecords("Account_Status_Change.IMD") Then
	Set db = Client.OpenDatabase("Cust Txn Status Change.IMD")
	Set task = db.VisualConnector
		id0 = task.AddDatabase ("Cust Txn Status Change.IMD")
		id1 = task.AddDatabase ("Account_Status_Change.IMD")
		task.MasterDatabase = id0
		task.AppendDatabaseNames = FALSE
		task.IncludeAllPrimaryRecords = TRUE
		task.AddRelation id0, "ACID", id1, "TABLE_KEY"
		task.AddFieldToInclude id0, "TRANSACTION_DATE"
		task.AddFieldToInclude id0, "TRAN_ID"
		task.AddFieldToInclude id0, "ACID"
		task.AddFieldToInclude id0, "TRANSACTION_AMT"
		task.AddFieldToInclude id0, "TRAN_PARTICULAR"
		task.AddFieldToInclude id0, "ENTRY_USER_ID"
		task.AddFieldToInclude id0, "PSTD_USER_ID"
		task.AddFieldToInclude id0, "RCRE_USER_ID"
		task.AddFieldToInclude id0, "RCRE_TIME_DATE"
		task.AddFieldToInclude id0, "RCRE_TIME_TIME"
		task.AddFieldToInclude id0, "CUST_ID"
		task.AddFieldToInclude id0, "TRANSACTION_CRNCY_CODE"
		task.AddFieldToInclude id0, "PART_TRANSACTION_TYPE"
		task.AddFieldToInclude id1, "REF_NUM"
		task.AddFieldToInclude id1, "TABLE_NAME"
		task.AddFieldToInclude id1, "TABLE_KEY"
		task.AddFieldToInclude id1, "ACID"
		task.AddFieldToInclude id1, "ENTERER_ID"
		task.AddFieldToInclude id1, "AUTH_ID"
		task.AddFieldToInclude id1, "MODIFIED_FIELDS_DATA"
		task.AddFieldToInclude id1, "AUDIT_DATE_DATE"
		task.AddFieldToInclude id1, "AUDIT_DATE_TIME"
		task.AddFieldToInclude id1, "COMP_STATUS_CHANGE"
		task.AddFieldToInclude id1, "COMP_CHANGE_FROM"
		task.AddFieldToInclude id1, "COMP_CHANGE_TO"
		task.CreateVirtualDatabase = False
		dbName = "Transaction Audit Change Int.IMD"
		task.OutputDatabaseName = dbName
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function
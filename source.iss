Function EJoinExtResult_Summ
If haveRecords("TM5_DTLS.IMD") Then
	Set db = Client.OpenDatabase("TM5_DTLS.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "POST_DATE"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "BRANCH_NAME"	
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.Criteria = ""
	dbName = "TM5_Summ1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function FRename_ResultFields
If haveRecords("TM5_SUMM1.IMD") Then
	Set db = Client.OpenDatabase("TM5_SUMM1.IMD")

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_TRANSACTIONS"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS1", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_SUMMARY"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
	task.PerformTask
End If
End Function


Function ASummHist_Average
If haveRecords("Tran_Hist_Average.IMD") Then
	Set db = Client.OpenDatabase("Tran_Hist_Average.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "TRANSACTION_TYPE"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.Criteria = ""
	dbName = "Summ_Hist_Average.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function
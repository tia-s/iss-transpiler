' Analysis: Summarization
Function Summarization
	Set db = Client.OpenDatabase("Append Databases.IMD")
	Set task = db.Summarization
	task.UseQuickSummarization = TRUE
	task.AddFieldToSummarize "CU"
	task.AddFieldToInc "MB_NUM"
	task.AddFieldToInc "LAST_NAME"
	task.AddFieldToInc "MIDDLE_NAME"
	task.AddFieldToTotal "ZZZ_MC_LIMIT"
	task.AddFieldToTotal "XFERRED_FL"
	dbName = "Summarization2.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = TRUE
	task.ResultName = "Summarization"
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
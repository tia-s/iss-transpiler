' Join to get File to apply criteria to get results.
Function DExtResults_INT
If haveRecords("Summ_Tran_Today.IMD") Then
Set db = Client.OpenDatabase("Summ_Tran_Today.IMD")
	Set task = db.JoinDatabase
   If haveRecords("Summ_Hist_Average.IMD") Then
	task.FileToJoin "Summ_Hist_Average.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TRANSACTION_AMOUNT_AVERAGE"
	task.AddSFieldToInc "TRANSACTION_AMOUNT_SUM"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.AddMatchKey "TRANSACTION_TYPE", "TRANSACTION_TYPE", "A"
	task.CreateVirtualDatabase = False
	dbName = "TM1_INT.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function
	
Function BExtResult_Summ
If haveRecords("TM5_Summ_INT.IMD") Then
	Set db = Client.OpenDatabase("TM5_Summ_INT.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "TM5_SUMM_INT2.IMD"
	task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_SUM >= " & e_TM5_Thresh 
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function
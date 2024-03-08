
' Analysis: Duplicate Key Detection
Function DuplicateKeyDetection
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.DupKeyDetection
	task.AddFieldToInc "MAIL_COUNTRY_CD"
	task.AddFieldToInc "GUID_ID"
	task.AddKey "MB_NUM", "A"
	task.OutputDuplicates = TRUE
	dbName = "Duplicate.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
End Function
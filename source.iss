' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Append Databases.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "VVA_INT1"
	task.AddFieldToInc "TAX_EXEMPT_AMT"
	task.AddField "TEST", "", WI_BOOL, 1, 0, "0"
	dbName = "EXTRACTION7.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.AddExtraction "EXTRACTION8.IMD", "", ""
	task.PerformTask 1, 2
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
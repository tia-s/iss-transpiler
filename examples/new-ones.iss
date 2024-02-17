
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
	Client.OpenDatabase (dbName)
End Function

' Analysis: Duplicate Key Detection
Function DuplicateKeyDetection
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.DupKeyDetection
	task.AddFieldToInc "VVA_INT3"
	task.AddFieldToInc "MAIL_PROVINCE_CD"
	task.AddKey "MB_NUM", "D"
	task.OutputDuplicates = FALSE
	dbName = "Non-duplicate.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

NB: can only have one key but can add multiple fields

' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "EXTRACTION2.IMD"
	dbName = "Append Databases.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Analysis: Fuzzy Duplicate
Function FuzzyDuplicate
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.FuzzyDuplicate
	task.AddFieldToInc "VVA_INT3"
	task.AddMatchField "BRANCH"
	dbName = "Fuzzy Duplicate.IMD"
	task.OutputDBName = dbName
	task.CreateVirtualDatabase = False
	task.AllowRecordsInMultipleFuzzyGroups = True
	task.IncludeExactDuplicates = True
	task.MatchCase = False
	task.SimilarityDegreeThreshold = 0.8
	task.OutputType = WI_FD_OUTPUT_MATCHES
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

NB: there are other fuzzy duplicates





' File: Compare Databases
Function CompareDatabase
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.CompareDB
	task.AddMatchKey "CU", "MB_NUM", "A"
	dbName = "Compare Databases.IMD"
	task.PerformTask dbName, "", "CU", "MB_NUM", "PUB_PART2.IMD"
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Top Records Extraction
Function TopNExtraction
	Set db = Client.OpenDatabase("IndExt.IMD")
	Set task = db.TopRecordsExtraction
	task.AddFieldToInc "MIDDLE_NAME"
	task.AddKey "MIDDLE_NAME", "A"
	task.AddKey "ADDR3", "D"
	dbName = "TopRecordsExtraction2.IMD"
	task.OutputFileName = dbName
	task.NumberOfRecordsToExtract = 1
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Data: Key Value Extraction
Function KeyValueExtraction
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.KeyValueExtraction
	dim myArray(0,0)
	myArray(0,0) = "51604110"
	task.AddFieldToInc "ADDR2"
	task.AddKey "MB_NUM", "D"
	task.DBPrefix = "KeyVal"
	task.CreateMultipleDatabases = TRUE
	task.CreateVirtualDatabase = False
	task.ValuesToExtract myArray
	task.PerformTask
	dbName = task.DBName
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase(dbName)
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "EXTRACTION6.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, 24
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

NB: perform task specifies how many rows to include in the extraction

' Data: Index Database
Function IndexDatabase
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.Index
	task.AddKey "CU", "A"
	task.Index FALSE
	Set task = Nothing
	Set db = Nothing
End Function
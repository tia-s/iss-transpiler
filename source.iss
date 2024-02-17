' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("PUB.mb.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "EXTRACTION2.IMD"
	dbName = "Append Databases.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
End Function
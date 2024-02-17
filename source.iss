' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Join Databases1.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "PUB.mb.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "GROUP_ID_AVERAGE", "CU", "A"
	task.CreateVirtualDatabase = False
	dbName = "Join Databases2.IMD"
	task.PerformTask dbName, "", WI_JOIN_NOC_PRI_MATCH
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
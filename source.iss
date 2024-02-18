Function BAJoinCustomers
If haveRecords("dbo.Registrations.IMD") Then
Set db = Client.OpenDatabase("dbo.Registrations.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "REPORTSVR.VDPUCID01.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ADDRESS1"
	task.AddSFieldToInc "ADDRESS2"
	task.AddSFieldToInc "ADDRESS3"
	task.AddSFieldToInc "BIRTHDAT"
	task.AddSFieldToInc "HOLDRTYP"
	task.AddSFieldToInc "FIRSTNAME"
	task.AddSFieldToInc "SURNAME"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Registrations.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function
' Append Field
Function AddFields

If haveRecords("REPORTSVR.VDPUTR04N.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR04N.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = ""
	field.Length = 15
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 	

If haveRecords("REPORTSVR.VDPUTR081.IMD") Then	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR081.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = ""
	field.Length = 12
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 

If haveRecords("REPORTSVR.VDPUTR021IMD") Then	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR021.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = ""
	field.Length = 15
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If

If haveRecords("REPORTSVR.VDPUTR011.IMD") Then	
	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR011.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = ""
	field.Length = 12
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If

If haveRecords("REPORTSVR.VDPUCID01.IMD") Then	
	
	
		Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BIRTHDAT"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "BIRTHDAT_DATE", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 
End Function

Function AExt_Master
If haveRecords("REPORTSVR.VDPUTP01.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTP01.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "ACCTNOP"
	dbName = "Master_Funds.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function BACreateDate
If haveRecords("REPORTSVR.VDPUCID01.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE_CREATED"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = ""
	task.AppendField field
	task.PerformTask
End If
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function
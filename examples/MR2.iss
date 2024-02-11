Function A_CreateYearMonth
If haveRecords("REPORTSVR.VDPUCID01.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "YEARMONTH"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@MID(@Dtoc(DATE_CREATED, ""YYYYMMDD""), 1, 6)"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
	Set field = Nothing
End Function


Function O_Rename_Fields
If haveRecords("MR2.IMD") Then
Set db = Client.OpenDatabase("MR2.IMD")
	Set task  = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE_OF_BIRTH"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "BIRTHDAT"
	task.AppendField field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "HOLDER_TYPE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@CompIf(HOLDRTYP == ""C"",""CORPORATION"", HOLDRTYP == ""I"",""INDIVIDUAL"",1,"""")"
	field.Length = 15
	task.AppendField field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "OCCUPATION"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 30
	task.AppendField field
	task.PerformTask
 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "POWEROFATTORNEYEXPIRE_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(POWEROFATTORNEYEXPIRE,""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask
	
			
	Set task = Nothing
	Set db = Nothing
End if
	Set field = Nothing
End Function


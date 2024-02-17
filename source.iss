Function G_Rename_SummFields
If haveRecords("TM12_SUMM.IMD") Then
Set db = Client.OpenDatabase("TM12_SUMM.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_TRANSACTIONS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS_SUM", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_DAYS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS1", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_SUMMARY"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_SUM_SUM", field
	task.PerformTask

	Set task  = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "STAFF"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = ""
	field.Length = 50
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If
End Function
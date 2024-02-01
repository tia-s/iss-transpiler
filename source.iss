' Create a many to many match of the records to be analysed.
Function Join_Transaction_Status
If haveRecords("Transaction Audit Change.IMD") Then
	Set db = Client.OpenDatabase("Transaction Audit Change.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_DORMANT_TXN1"
		task.AppendField field
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
If haveRecords("Transaction Audit Change.IMD") Then
	Set db = Client.OpenDatabase("Transaction Audit Change.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_DORMANT_TXN1"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = ""
		field.Length = 1
		task.AppendField field
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function
' Create a many to many match of the records to be analysed.


Function Get_Final_Cheque_File

	Client.CloseAll
	
	'For handling of empty files:
	Dim foreign_has_records As Boolean
	Dim regular_has_records As Boolean
	Dim clearing_has_records As Boolean

	Dim file_index As Integer			'Used as an array index
	Dim file_counter As Integer		'Count the occurrences of non-empty files
	Dim counter As Integer			'A loop counter
	Dim append_index As Integer		'Used as an array index
	Dim database_to_append(4) As String	'Array to hold the names of databases to append 

	'Initialise variables
	file_index = 1
	append_index = 1
	file_counter = 0

	'Check each file for existence of records
	foreign_has_records = ContainsRecords("Foreign Cheque Details.IMD")
	If foreign_has_records Then

		file_counter = file_counter + 1
		database_to_append(file_index) = "Foreign Cheque Details.IMD"
		file_index = file_index + 1
		
	End If

	regular_has_records = ContainsRecords("Regular Cheque.IMD")
	If  regular_has_records Then

		file_counter = file_counter + 1
		database_to_append(file_index) = "Regular Cheque.IMD"	
		file_index = file_index + 1
		
	End If

	clearing_has_records = ContainsRecords("Clearing Cheques.IMD") 		
	If clearing_has_records Then
	
		file_counter = file_counter + 1	
		database_to_append(file_index) = "Clearing Cheques.IMD"		
		file_index = file_index + 1
	End If


	'Start the append
	
	If file_counter = 0 Then						' If no files exist
	Exit Sub
	
	ElseIf file_counter = 1 Then					'If only one file, rename that file to the name of the final append file
	
	                Client.CloseAll
	                Set ProjectManagement = client.ProjectManagement
	                ProjectManagement.RenameDatabase database_to_append(1), "All Cheques"
	                Set ProjectManagement = Nothing
		Set task = Nothing
		Set db = Nothing
	
	Else								'If more than one, append the databases that are not empty
	
		Set db = Client.OpenDatabase(database_to_append(1))	'The starting database
		Set task = db.AppendDatabase
	
			'Add the other databases
			For counter = 1 To (file_counter-1)
				append_index = append_index + 1
				task.AddDatabase (database_to_append(append_index))
			Next counter
		
									'Final database name
		dbName = "All Cheques.IMD"
			
		task.PerformTask dbName, ""
		Set task = Nothing
		Set db = Nothing

	End If

End Function

---



'Total wires in by other party and party.
Function SummarizeInWireOthPrty
If haveRecords("Inward Wire Cust Acct Branch.IMD") Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "USD"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """USD"""
		field.Length = 3
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
		
	Set field = db.TableDef.NewField
		field.Name = "RUN_DATE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@Dtoc(@Date(),""YYYY-MM-DD"")"
		field.Length = 10
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
	task.AppendField field
	
	task.PerformTask
	Set field = Nothing	
	Set task = Nothing
	Set db = Nothing
End If 	

---

' File - Export Database: MDB2000
Function ExportSummaryandDetail

If haveRecords("WT6_Summ.IMD") Then
	Set db = Client.OpenDatabase("WT6_Summ.IMD")
	Set task = db.ExportDatabase
	 task.IncludeAllFields
	eqn = ""
	task.PerformTask "Reports\WT6_Multiple Wire Draft Multiple Branch Summary.MDB", "MultiWireDrftMultiBranchSumm", "MDB2000", 1, db.Count, eqn
RESULTSLOG(db.name)
		Set task = Nothing
	Set db = Nothing
	
Else 
NORESULTSLOG("WT6_SUMMARY.IMD") 
End If
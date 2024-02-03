'====================================================================================================
'	Test#: 		WT13 - Identify Sequential drafts to multi persons same cust
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	31/03/2014 02:11:25 PM
'
'	KWHYTE	May 06, 2015: Modified to change name of output MDB file
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "WT13 - Identify Sequential drafts to multi persons same cust"
Const scriptname_log ="WT13 - Identify Sequential drafts to multi persons same cust.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main

	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	'Client.CloseAll
	CleanUp
 	
	
	Call Get_Customer_Details
	Call Get_Multiple_Draft_Same_Customer_Diff_Recipient
	Call Get_Sequential_Drafts

	Call Get_Multiple_Draft_Same_Customer_Diff_Recipient_Summary
	Call Get_Multiple_Draft_Same_Customer_Diff_Recipient_Details
		
	Call Join_RiskRating()
	
	Call Export_to_MDB("WT13_Multi Drafts Summ", "")
	Call Export_to_MDB("WT13_Multi Draft Details", "")
		
	finalRoutine:
	
	If err.description <> "" Or err.number <> 0 Then
		message  = err.description & " " & err.number
		info = "Error"
	Else
		message = "Script Completed Successfully"
		info = "Information"
	End If	
	
	Call logfile(scriptname_log, "End", "Data Analysis", info, message & errors_string)


	Client.CloseAll
	CleanUp
	'Client.Quit
End Sub


'---------------------------------------------------------------------------------------------------------------------------------------
' Logfile(ByVal Log_Step As String, ByVal Log_Step As String, ByVal Log_Action As String, ByVal Log_MsgType As String, ByVal Log_Message As String)
'
' Input:
' filename_log		- {String} Which file the analysis is on
' Log_Step		- {String} Which step is being run
' Log Action		- {String} Which Action is performed
' Log_Msgtype		- {String} Log Type (Informational, Error, Warning)
' Log_Message		- {String} Log Message
'
' Returns: 		Nothing
'
' Description: This function creates and appends to a logfile 
'---------------------------------------------------------------------------------------------------------------------------------------	
Function logfile(ByVal filename_log As String,  ByVal Log_Step As String, ByVal Log_Action As String, ByVal Log_MsgType As String, ByVal Log_Message As String)

On Error GoTo exit_logfile

If e_debug <> True Then Exit Sub

Dim logfilename As String
Dim newtable As Object
Dim addedfield As Object
Dim db1 As Object
Dim rs1 As Object
Dim rec1 As Object
Dim tbb As Object
Dim fields As Double
Dim i As Double
Dim field As Object
Dim sdir As String


If (Len(e_logfilename) > 0) Then logfilename = e_logfilename & ".imd" Else logfilename = "log_file.imd"

	'Create the table if it doesn't exist. 
	Set pm = Client.ProjectManagement
	If Not pm.DoesDatabaseExist(logfilename) Then
		Set NewTable = Client.NewTableDef
		Set AddedField = NewTable.NewField
		AddedField.Name = "LOG_DATE"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = "LOG_TIME"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False				
		AddedField.Name = "FILENAME"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = "SCRIPTNAME"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=50	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = "SCRIPTSTEP"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=50	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = "ACTION"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = "MSG_TYPE"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False		
		AddedField.Name = "MESSAGE"
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		Set db1 = Client.NewDatabase(Logfilename, "", NewTable)
		db1.commitdatabase
		db1.close
		Set db1 = Nothing
		Set addedfield = Nothing
		Set newtable = Nothing
	End If

	
	'Write the log message	
	Set db1 = Client.OpenDatabase(Logfilename)
	Set rs1 =  db1.RecordSet
	Set rec1 = rs1.NewRecord
	Set tbb = db1.tabledef
	fields = tbb.count
	For i = 1 To fields
		Set field =tbb.getfieldat(i)	
		field.protected = false
	Next i

		rec1.setcharvalueat 1, Format (Now(), "Short Date")
		rec1.setcharvalueat 2, Format (Now(), "Medium Time")
		If filename_log <> "" Then 	rec1.setcharvalueat 3, filename_log
		If scriptname_log <> "" Then rec1.setcharvalueat 4, scriptname_log
		If Log_Step  <> "" Then rec1.setcharvalueat 5, Log_Step   
		If Log_Action <> "" Then rec1.setcharvalueat 6, Log_Action
		If Log_MsgType <> "" Then rec1.setcharvalueat 7, Log_MsgType
		If Log_Message <> "" Then rec1.setcharvalueat 8, Log_Message
		
		rs1.appendrecord rec1
	For i = 1 To fields
		Set field =tbb.getfieldat(i)	
		field.protected = true
	Next i
	db1.commitdatabase
	db1.close
	Set field = Nothing
	Set tbb = Nothing
	Set rec1 = Nothing
	Set rs1 = Nothing
	Set db1 = Nothing
	
exit_logfile:	
	
End Function

'---------------------------------------------------------------------------------------------------------------------------------------
Function haveRecords(ByVal dbName As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------
	Dim records As Double
	Dim db As Object
	Dim rs As Object
	Dim pm As ProjectManagement
	records = 0
	haveRecords = False
	Set pm = Client.ProjectManagement
	If pm.DoesDatabaseExist(dbName) Then
		Set db = Client.OpenDatabase(dbName)
			Set rs =  db.RecordSet
				If rs.count > 0 Then
					haveRecords = True
				Else
					errors_string = errors_string & " with errors -" & dbname & " has no records." & Chr(10)
					Call logfile(dbname, division, "haverecords", "Error", "Database does not have records.")
				End If
			Set rs = Nothing
		db.close
		Set db = Nothing
	Else
		errors_string = errors_string & " with errors -" & dbname & " missing." & Chr(10)
		Call logfile(dbname, division, "haverecords", "Error", "Database does not exist.")
		
	End If
End Function

Function FieldExist(ByVal dbname As String, ByVal fieldname As String) As Boolean
FieldExist = False

Dim a_count As Double
Dim db As Object
Dim table As Object
Dim fields As Double
Dim cnfield As Object

	Set db = Client.OpenDatabase(dbname)
		Set table = db.TableDef
		fields = table.count
		
		For a_count = 1 To fields
			Set cnfield = table.GetFieldat(a_count)
			If UCase(Trim(cnfield.name)) =  UCase(Trim(fieldname)) Then 
			                FieldExist = True
			                a_count = fields
			End If
		Next a_count
			
		Set cnfield = Nothing
		Set table = Nothing
		Set db = Nothing
End Function

Function Delete_Virtual_Field(TableName As String, Fieldname As String)
	Dim task As Object
	Dim db As Object
	Dim table As Object
	
	Set db = Client.OpenDatabase(TableName)
	                Set task = db.TableManagement
	              		  Set table = db.TableDef
				task.RemoveField Fieldname
				task.PerformTask
	                	Set task = Nothing
		Set table = Nothing
	Set db = Nothing
End Function

Function DeleteFile(NameOfFile As String)
	Client.CloseAll
	If fso.FileExists(Client.WorkingDirectory & Trim(NameOfFile)) = True Then Kill(Client.WorkingDirectory & Trim(NameOfFile))
End Function


'--------------------------------------------------------------------------------------
' get customer and Account information
'--------------------------------------------------------------------------------------
Function Get_Customer_Details

If haveRecords("OutGoing_Drafts_Transactions.IMD")  And haveRecords("General Acct Master Lite.IMD") Then
	Set db = Client.OpenDatabase("OutGoing_Drafts_Transactions.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "General Acct Master Lite.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "FORACID"
	task.AddSFieldToInc "CUST_ID"
	task.AddSFieldToInc "SOL_ID"
	task.AddSFieldToInc "SCHM_DESC"
	task.AddSFieldToInc "ACCT_OWNERSHIP"
	task.AddMatchKey "OPER_ACID", "ACID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Outward Drafts with Account Details.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("Outward Drafts with Account Details.IMD") And haveRecords("Active Customers.IMD") Then
	Set db = Client.OpenDatabase("Outward Drafts with Account Details.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Active Customers.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CUST_ID"
	task.AddSFieldToInc "CUST_NAME"
	task.AddSFieldToInc "CUST_PERM_ADDR1"
	task.AddSFieldToInc "CUST_PERM_ADDR2"
	task.AddSFieldToInc "CUSTOMER_TYPE"
	task.AddMatchKey "CUST_ID", "CUST_ID", "A"
	task.Criteria = "ACCT_OWNERSHIP <> ""O"""
	task.CreateVirtualDatabase = False
	dbName = "Outward Drafts with Acct Cust Details.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End if
		
End Function

Function Get_Multiple_Draft_Same_Customer_Diff_Recipient

If haveRecords("Outward Drafts with Acct Cust Details.IMD") Then
	Set db = Client.OpenDatabase("Outward Drafts with Acct Cust Details.IMD")
	Set task = db.DupKeyExclusion
	task.IncludeAllFields
	task.AddKey "CUST_ID", "A"
	task.AddKey "RCRE_TIME_DATE", "A"
	task.DifferentField = "OTHER_PARTY_NAME"
	dbName = "Multi Draft same Cust Diff Recipient.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
End If

End Function

Function Get_Sequential_Drafts

If haveRecords("Multi Draft same Cust Diff Recipient.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "BILL_NUMBER"
				field.Description = ""
				field.Type = WI_VIRT_NUM
				field.Equation = "@val(@right(BILL_ID, 5))"
				field.Length = 6
			task.AppendField field
			task.PerformTask
		Set field = Nothing
		Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("Multi Draft same Cust Diff Recipient.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient.IMD")
		Set task = db.Sort
			task.AddKey "CUST_ID", "A"
			task.AddKey "BILL_ID", "A"
			dbName = "Multi Draft same Cust Diff Recipient Sorted.IMD"
			task.PerformTask dbName
		Set task = Nothing
	Set db = Nothing
End if		
	
If haveRecords("Multi Draft same Cust Diff Recipient Sorted.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient Sorted.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "SEQUENTIAL"
				field.Description = ""
				field.Type = WI_VIRT_NUM
				field.Equation = "@If(BILL_NUMBER == (@GetPreviousValue(""BILL_NUMBER"") + 1) .OR. (BILL_NUMBER + 1) == @GetNextValue(""BILL_NUMBER""), 1, 0)"
				field.Decimals = 0
				task.AppendField field
				task.PerformTask
			Set field = Nothing
		Set task = Nothing
	Set db = Nothing	
End if
	
If haveRecords("Multi Draft same Cust Diff Recipient Sorted.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient Sorted.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Multi Draft same Cust Diff Recipient Seq Drafts.IMD"
	task.AddExtraction dbName, "", "SEQUENTIAL==1"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If

	
End Function

Function Get_Multiple_Draft_Same_Customer_Diff_Recipient_Summary

If haveRecords("Multi Draft same Cust Diff Recipient Seq Drafts.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient Seq Drafts.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "CUST_ID"
	task.AddFieldToInc "CUST_NAME"
	task.AddFieldToInc "CUST_PERM_ADDR1"
	task.AddFieldToInc "CUST_PERM_ADDR2"
	task.AddFieldToInc "CUSTOMER_TYPE"
	task.AddFieldToInc "SOL_ID1"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_JM"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_US"
	dbName = "Multi Drafts Same Cust Diff Recipient Summ1.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Multi Drafts Same Cust Diff Recipient Summ1.IMD") Then
	Set db = Client.OpenDatabase("Multi Drafts Same Cust Diff Recipient Summ1.IMD")
	Set task = db.JoinDatabase
   If haveRecords("NCB_BRANCHES.IMD") Then
   	task.FileToJoin "NCB_BRANCHES.IMD"
   	task.IncludeAllPFields
   	task.AddSFieldToInc "BR_NAME"
   	task.AddMatchKey "SOL_ID1", "MICR_BRANCH_CODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "Multi Drafts Same Cust Diff Recipient Summ2.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
End If
	Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("Multi Drafts Same Cust Diff Recipient Summ2.IMD") Then
	Set db = Client.OpenDatabase("Multi Drafts Same Cust Diff Recipient Summ2.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "CUST_ID"
	task.AddFieldToInc "CUST_NAME"
	task.AddFieldToInc "CUST_PERM_ADDR1"
	task.AddFieldToInc "CUST_PERM_ADDR2"
	task.AddFieldToInc "CUSTOMER_TYPE"
	task.AddFieldToInc "SOL_ID1"
	task.AddFieldToInc "BR_NAME"
	task.AddFieldToInc "NO_OF_RECS"
	task.AddFieldToInc "TRANSACTION_AMOUNT_JM_SUM"
	task.AddFieldToInc "TRANSACTION_AMOUNT_US_SUM"
	dbName = "WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD"
	'task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US_SUM >= "& e_TTR_Value_US &" .AND. NO_OF_RECS >="& e_Different_Persons_Count
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	

If haveRecords("WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_DIFFERENT_PERSONS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	
	Set field = db.TableDef.NewField
	field.Name = "TOTAL_TRANSACTION_AMOUNT_JM"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_JM_SUM", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_BRANCH_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 30
	task.ReplaceField "BR_NAME", field
	
	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_SOL_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 5
	task.ReplaceField "SOL_ID1", field

	Set field = db.TableDef.NewField
	field.Name = "TOTAL_TRANSACTION_AMOUNT_US"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "TRANSACTION_AMOUNT_US_SUM", field
	
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "CUST_NAME", field
	
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ADDRESS1"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 45
	task.ReplaceField "CUST_PERM_ADDR1", field	
	
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ADDRESS2"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 45
	task.ReplaceField "CUST_PERM_ADDR2", field
	
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "CUST_ID", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD") And haveRecords("Customer Master Lite.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Customer Master Lite.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "PRIMARY_SOL_ID"
	task.AddMatchKey "CUSTOMER_ID", "CUST_ID", "A"
	task.CreateVirtualDatabase = False
	dbName = "WT13_Multi Drafts Same Cust Diff Recipient Summ4.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("WT13_Multi Drafts Same Cust Diff Recipient Summ4.IMD") And haveRecords("NCB_BRANCHES.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Drafts Same Cust Diff Recipient Summ4.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "NCB_BRANCHES.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "BR_NAME"
	task.AddMatchKey "PRIMARY_SOL_ID", "MICR_BRANCH_CODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "WT13_Multi Drafts Summ_tmp.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("WT13_Multi Drafts Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Drafts Summ_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_BRANCH_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "BR_NAME", field
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
End Function

' File: Join Databases
Function Get_Multiple_Draft_Same_Customer_Diff_Recipient_Details

If haveRecords("Multi Draft same Cust Diff Recipient Seq Drafts.IMD") And haveRecords("WT13_Multi Drafts Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("Multi Draft same Cust Diff Recipient Seq Drafts.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "WT13_Multi Drafts Summ_tmp.IMD"
	task.AddSFieldToInc "CUSTOMER_ID"
	task.AddSFieldToInc "CUSTOMER_NAME"
	task.AddSFieldToInc "CUSTOMER_ADDRESS1"
	task.AddSFieldToInc "CUSTOMER_ADDRESS2"
	task.AddSFieldToInc "CUSTOMER_TYPE"
	task.AddSFieldToInc "ACCOUNT_SOL_ID"
	task.AddSFieldToInc "ACCOUNT_BRANCH_NAME"
	task.AddPFieldToInc "FORACID"	
	task.AddPFieldToInc "SCHM_DESC"
	task.AddPFieldToInc "BILL_ID"
	task.AddPFieldToInc "TRANSFER_DATE_DATE"
	task.AddPFieldToInc "TRANSACTION_DATE"
	task.AddPFieldToInc "BILL_CNTRY_CODE"
	task.AddPFieldToInc "TRANSACTION_AMT"
	task.AddPFieldToInc "TRANSACTION_AMOUNT_JM"
	task.AddPFieldToInc "TRANSACTION_AMOUNT_US"
	task.AddPFieldToInc "OTHER_PARTY_NAME"
	task.AddPFieldToInc "OTHER_PARTY_ADDR_1"
	task.AddPFieldToInc "OTHER_PARTY_ADDR_2"
	task.AddPFieldToInc "OTHER_PARTY_ADDR_3"
	task.AddPFieldToInc "OTHER_PARTY_CNTRY_CODE"
	task.AddMatchKey "CUST_ID", "CUSTOMER_ID", "A"
	task.CreateVirtualDatabase = False
	dbName = "WT13_Multi Draft Same Cust Diff Recipient Details Int.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("TBAADM.FEX_CLEAN_INST_TABLE.IMD") And haveRecords("WT13_Multi Draft Same Cust Diff Recipient Details Int.IMD") Then	
	Set db = Client.OpenDatabase("TBAADM.FEX_CLEAN_INST_TABLE.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "WT13_Multi Draft Same Cust Diff Recipient Details Int.IMD"
			task.AddPFieldToInc "INSTRMNT_NUM"
			task.AddPFieldToInc "INSTRMNT_TYPE"
			task.AddPFieldToInc "INSTRMNT_AMT"
			task.AddPFieldToInc "INSTRMNT_DATE_DATE"
			task.IncludeAllSFields
			task.AddMatchKey "BILL_ID", "BILL_ID", "A"
			task.CreateVirtualDatabase = False
			task.Criteria = "DEL_FLG == ""N"""
			dbName = "WT13_Multi Draft Details.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If

If haveRecords("WT13_Multi Draft Details.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Draft Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_NUMBER"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "FORACID", field
	
	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_DESC"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "SCHM_DESC", field
	
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_CNTRY_CODE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 5
	task.ReplaceField "BILL_CNTRY_CODE", field
	
	Set field = db.TableDef.NewField
	field.Name = "OTHER_PARTY_COUNTRY"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 5
	task.ReplaceField "OTHER_PARTY_CNTRY_CODE", field
	
	Set field = db.TableDef.NewField
	field.Name = "INSTRUMENT_NUM"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "INSTRMNT_NUM", field

	Set field = db.TableDef.NewField
	field.Name = "INSTRUMENT_TYPE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 5
	task.ReplaceField "INSTRMNT_TYPE", field

	Set field = db.TableDef.NewField
	field.Name = "INSTRUMENT_AMT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "INSTRMNT_AMT", field
	
	Set field = db.TableDef.NewField
	field.Name = "INSTRUMENT_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "INSTRMNT_DATE_DATE", field
	
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
End if

End Function

Function Join_RiskRating
If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("WT13_Multi Drafts Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("WT13_Multi Drafts Summ_tmp.IMD")	
	Set task = db.JoinDatabase
task.FileToJoin "Customer_Turnover_wRisk.IMD" 
		task.IncludeAllPFields
		task.AddSFieldToInc "MONTHLY_DEPOSIT"
		task.AddSFieldToInc "RISK_SCORE"
		task.AddSFieldToInc "RISK_CATEGORY"
		task.AddSFieldToInc "OCCUPATION"
		task.AddSFieldToInc "INDUSTRY"
		task.AddSFieldToInc "ANNUAL_INCOME"
		task.AddMatchKey "CUSTOMER_ID", "CUST_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "WT13_Multi Drafts Summ.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Export_to_MDB(Filename As String, eqn As String)

If haveRecords(Trim(Filename)+".IMD") Then
	Set db = Client.OpenDatabase(Trim(Filename)+".IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	task.PerformTask "Reports\"&Trim(Filename)&".MDB", "Database", "MDB2000", 1, db.Count, eqn
	RESULTSLOG(db.name)
	Set task = Nothing
	Set db = Nothing
Else 
NORESULTSLOG("WT13_SUMMARY.IMD") 
End If
End Function


Function CleanUp
	DeleteFile("Multi Draft same Cust Diff Recipient.IMD")
	DeleteFile("Outward Drafts with Account Details.IMD") 
	DeleteFile("Outward Drafts with Acct Cust Details.IMD")
	DeleteFile("Multi Draft same Cust Diff Recipient Sorted.IMD")
	DeleteFile("Multi Draft same Cust Diff Recipient Seq Drafts.IMD")
	DeleteFile("Multi Drafts Same Cust Diff Recipient Summ.IMD")
	DeleteFile("WT13_Multi Draft Details.IMD")
	'DeleteFile("WT13_Multi Drafts Summ.IMD")
	DeleteFile("WT13_Multi Drafts Summ_tmp.IMD")
	DeleteFile("Multi Drafts Same Cust Diff Recipient Summ1.IMD")
	DeleteFile("WT13_Multi Draft Same Cust Diff Recipient Details Int.IMD")	
	DeleteFile("Multi Drafts Same Cust Diff Recipient Summ2.IMD")
	DeleteFile("Multi Drafts Same Cust Diff Recipient Summ3.IMD")
	DeleteFile("Multi Drafts Same Cust Diff Recipient Summ4.IMD")
	DeleteFile("WT13_Multi Drafts Same Cust Diff Recipient Summ4.IMD")
	DeleteFile("WT13_Multi Drafts Same Cust Diff Recipient Summ3.IMD")
End Function




Function RESULTSLOG(FileName1 As String)

On Error GoTo err_handle

Dim LogCountFile As TextStream
Dim LogName As String
Dim Path As String
Dim FileName As String
  
FileName= Right(FileName1,(Len(FileName1))-3-InStr(1, FileName1, "ory\") )
Set db = Client.OpenDatabase(FileName)
recnum = db.count
Path = Client.WorkingDirectory & "Reports\"
LogName = "ResultsWeekly.csv"

'Create the log if it does not exist and writes reader record
  If Not fso.FileExists(Path & LogName) Then
 
	fso.CreateTextFile (Path & LogName) 
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 2, True)
	LogCountFile.WriteLine("Log_Date" & Chr(9) & "ResultsFile_Name" & Chr(9) & "Record_Count")
	LogCountFile.Close
 
  End If 

' Writes records To file that already exists
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 8, True) 
	LogCountFile.WriteLine (Now & Chr(9) & FileName & Chr(9) & recnum)
	LogCountFile.Close
	Set db = Nothing
err_handle:
    Client.CloseAll
End Function


 Function NORESULTSLOG(FileName As String)

On Error GoTo err_handle

Dim LogCountFile As TextStream
Dim LogName As String
Dim Path As String

recnum = 0
Path = Client.WorkingDirectory & "Reports\"
LogName = "ResultsWeekly.csv"

'Create the log if it does not exist and writes reader record
  If Not fso.FileExists(Path & LogName) Then
 
	fso.CreateTextFile (Path & LogName) 
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 2, True)
	LogCountFile.WriteLine("Log_Date" & Chr(9) & "ResultsFile_Name" & Chr(9) & "Record_Count")
	LogCountFile.Close
 
  End If 

' Writes records To file that already exists
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 8, True) 
	LogCountFile.WriteLine (Now & Chr(9) & FileName & Chr(9) & recnum)
	LogCountFile.Close
	Set db = Nothing
err_handle:
    Client.CloseAll
End Function

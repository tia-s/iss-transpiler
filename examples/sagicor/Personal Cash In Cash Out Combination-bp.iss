'====================================================================================================
'	Test#: 		Personal Cash In Cash Out Combination
'	Risk:		Personal customers may be conducting cash deposit and withdrawals over the threshold.
' 	Objective:	Identify personal customers that deposit cash and withdraw cash where the total is over the threshold.
' 	Frequency:	FT
' 	Last Modified:	Jan-14-2015
'====================================================================================================
'	Script Dependencies: Import Interim
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "Personal Cash In Cash Out Combination"
Const scriptname_log ="Personal Cash In Cash Out Combination.iss"
Global errors_string As String
Const division = ""
Dim fso As Object

Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Client.CloseAll

	Z_CleanUp
	A_Get_CashInOut
	B_Get_First_CashIn_Per_Accont
	C_Join_FirstCashIn_To_AllCashInOut
	D_Get_First_CashOut_After_CashIn
	E_Get_CashIn_CashOut_Details
	F_ModifyFieldDetails
	G_GetSummaryRecords
	H_GetDetailRecords
	I_ModifyFieldSummary
	J_Get_Total_Credit_Debit
	K_Export

	
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
	Z_CleanUp
	Client.Quit
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


Function A_Get_CashInOut
' Get all Personal cash in and out for the fortnight.
	Set db = Client.OpenDatabase("FORTNIGHTLY_TRANSACTIONS.IMD")
	Set task = db.TableManagement
		Set field = db.TableDef.NewField
		field.Name = "RUN_DATE"
		field.Description = ""
		field.Type = WI_VIRT_DATE
		field.Equation = "@Date()"
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

If haveRecords("FORTNIGHTLY_TRANSACTIONS.IMD") Then
	Set db = Client.OpenDatabase("FORTNIGHTLY_TRANSACTIONS.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "SOURCE_ACCOUNT"
		task.AddFieldToInc "DEBIT_CREDIT"
		task.AddFieldToInc "CURRENCY"
		task.AddFieldToInc "OTHER_PARTY_ACC"
		task.AddFieldToInc "EXCLUDE_FROM_PROFILING"
		task.AddFieldToInc "DATE"
		task.AddFieldToInc "TRANSACTION_CODE"
		task.AddFieldToInc "AMOUNT"
		task.AddFieldToInc "ORIGINAL_AMOUNT"
		task.AddFieldToInc "REFERENCE"
		task.AddFieldToInc "DESCRIPTION"
		task.AddFieldToInc "FROM_CUSTOMER"
		task.AddFieldToInc "TO_ACCOUNT"
		task.AddFieldToInc "ROLE_TYPE"
		task.AddFieldToInc "CUSTOMER_NUMBER"
		task.AddFieldToInc "NAME"
		task.AddFieldToInc "COUNTRY_RESIDENCE"
		task.AddFieldToInc "COUNTRY_NATIONALITY"
		task.AddFieldToInc "DATEOFBIRTH"
		task.AddFieldToInc "ADDRESS_LINE_1"
		task.AddFieldToInc "ADDRESS_LINE_2"
		task.AddFieldToInc "ADDRESS_LINE_3"
		task.AddFieldToInc "ADDRESS_LINE_4"
		task.AddFieldToInc "ADDRESS_LINE_5"
		task.AddFieldToInc "GENDER"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "BUSINESS_CODE"
		task.AddFieldToInc "SSN"
		task.AddFieldToInc "COUNTRY_BIRTH"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "IDENTIFICATION_TYPE"
		task.AddFieldToInc "IDENTIFICATION_OTHER"
		task.AddFieldToInc "IDENTIFICATION_NUMBER"
		task.AddFieldToInc "ISSUING_AUTHORITY"
		task.AddFieldToInc "CONTACT_NUMBER"
		task.AddFieldToInc "COUNTRY_RISK"
		task.AddFieldToInc "CUSTOMER_STATUS"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "PEP"
		task.AddFieldToInc "ACCOUNT_TYPE"
		task.AddFieldToInc "STATUS"
		task.AddFieldToInc "ACCOUNT_NAME"
'		task.AddFieldToInc "CUS_NUMBER"
		task.AddFieldToInc "TRANSACTION_CURRENCY"
		task.AddFieldToInc "TRANSACTION_AMOUNT"
		task.AddFieldToInc "ORIGINAL_TRANSACTION_AMOUNT"
		task.AddFieldToInc "RUN_DATE"
		dbName = "Personal Cash In Out.IMD"
		task.AddExtraction dbName, "", "@Match(TRANSACTION_CODE, ""CASH"", ""ATMC"", ""DBTC"")  .AND. CUSTOMER_TYPE == ""P"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End if
End Function
	

' Get the first cash in generated per account
Function B_Get_First_CashIn_Per_Accont
If haveRecords("Personal Cash In Out.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out.IMD")
	Set task = db.TopRecordsExtraction
		task.AddFieldToInc "SOURCE_ACCOUNT"
		task.AddFieldToInc "DATE"
		task.AddKey "SOURCE_ACCOUNT", "A"
		task.AddKey "DATE", "A"
		task.Criteria = "DEBIT_CREDIT == ""C"""
		dbName = "First Personal Cash In.IMD"
		task.OutputFileName = dbName
		task.NumberOfRecordsToExtract = 1
		task.CreateVirtualDatabase = False
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Link cash In and out to place the first cash in date beside all the txns.
Function C_Join_FirstCashIn_To_AllCashInOut
If haveRecords("Personal Cash In Out.IMD") And haveRecords("First Personal Cash In.IMD")Then
	Set db = Client.OpenDatabase("Personal Cash In Out.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "First Personal Cash In.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "DATE"
		task.AddMatchKey "SOURCE_ACCOUNT", "SOURCE_ACCOUNT", "A"
		task.CreateVirtualDatabase = False
		dbName = "Personal Cash InOut First CashIn.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Get first Personal cash out after first cash in.
Function D_Get_First_CashOut_After_CashIn
If haveRecords("Personal Cash InOut First CashIn.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash InOut First CashIn.IMD")
	Set task = db.TopRecordsExtraction
		task.AddFieldToInc "SOURCE_ACCOUNT"
		task.AddFieldToInc "DATE"
		task.AddFieldToInc "DATE1"
		task.AddKey "SOURCE_ACCOUNT", "A"
		task.AddKey "DATE", "A"
		task.Criteria = "DEBIT_CREDIT == ""D"" .AND. DATE1 < Date"
		dbName = "Personal Cash Out After In.IMD"
		task.OutputFileName = dbName
		task.NumberOfRecordsToExtract = 1
		task.CreateVirtualDatabase = False
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Link Cash In and Out to cash out file
Function E_Get_CashIn_CashOut_Details
If haveRecords("Personal Cash In Out.IMD") And haveRecords("Personal Cash Out After In.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Personal Cash Out After In.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "SOURCE_ACCOUNT"
		task.AddSFieldToInc "DATE1"
		task.AddMatchKey "SOURCE_ACCOUNT", "SOURCE_ACCOUNT", "A"
		task.CreateVirtualDatabase = False
		dbName = "Personal Cash In Out Temp2.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End if
	
' Ensure ONLY records that fall after the first cash in are selected for the summary	
If haveRecords("Personal Cash In Out Temp2.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Temp2.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "Personal Cash In Out Temp.IMD"
		task.AddExtraction dbName, "", "DATE >= DATE1"
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing	
End if
End Function

Function F_ModifyFieldDetails
If haveRecords("Personal Cash In Out Temp.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Temp.IMD")
	Set task = db.TableManagement
		Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 50
		task.ReplaceField "SOURCE_ACCOUNT", field
	
		Set field = db.TableDef.NewField
		field.Name = "DIRECTION_OF_TRANSACTION"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 1
		task.ReplaceField "DEBIT_CREDIT", field

		Set field = db.TableDef.NewField
		field.Name = "ORIGINAL_CURRENCY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 5
		task.ReplaceField "CURRENCY", field

		Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 100
		task.ReplaceField "NAME", field

		Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "DATE", field

		Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_STATUS"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 50
		task.ReplaceField "STATUS", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function G_GetSummaryRecords
If haveRecords("Personal Cash In Out Temp.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Temp.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "ACCOUNT_TYPE"
		task.AddFieldToInc "ROLE_TYPE"
		task.AddFieldToInc "ACCOUNT_STATUS"
		task.AddFieldToInc "CUSTOMER_NUMBER"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "RUN_DATE"
		task.AddFieldToTotal "TRANSACTION_AMOUNT"
		task.AddFieldToInc "TRANSACTION_CURRENCY"
		dbName = "Personal Cash In Out Summ.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if

If haveRecords("Personal Cash In Out Summ.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Summ.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "Personal Cash In Out Summary.IMD"
		task.AddExtraction dbName, "", "@Between(TRANSACTION_AMOUNT_SUM," & e_Personal_Combine_Thresh - (e_Personal_combine_ratio * e_Personal_Combine_Thresh) & "," & e_Personal_Combine_Thresh + (e_Personal_combine_ratio * e_Personal_combine_Thresh) & ")"
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function H_GetDetailRecords
If haveRecords("Personal Cash In Out Temp.IMD") And haveRecords("Personal Cash In Out Summary.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Temp.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Personal Cash In Out Summary.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "RUN_DATE"
		task.AddMatchKey "ACCOUNT_NUMBER", "ACCOUNT_NUMBER", "A"
		task.CreateVirtualDatabase = False
		dbName = "Personal Cash In Out Details.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function I_ModifyFieldSummary
If haveRecords("Personal Cash In Out Summary.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Summary.IMD") 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_TRANSACTIONS"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("Personal Cash In Out Summary.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Summary.IMD") 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TOTAL_TRANSACTION_AMOUNT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function J_Get_Total_Credit_Debit
If haveRecords("Personal Cash In Out Details.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Details.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "ACCOUNT_NUMBER"
		task.AddFieldToTotal "TRANSACTION_AMOUNT"
		task.Criteria = "DIRECTION_OF_TRANSACTION ==""D"""
		dbName = "Debit Total.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("Personal Cash In Out Details.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Details.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "ACCOUNT_NUMBER"
		task.AddFieldToTotal "TRANSACTION_AMOUNT"
		task.Criteria = "DIRECTION_OF_TRANSACTION ==""C"""
		dbName = "Credit Total.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("Credit Total.IMD") Then
	Set db = Client.OpenDatabase("Credit Total.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "TOTAL_CREDIT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 2
		task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("Debit Total.IMD") Then
	Set db = Client.OpenDatabase("Debit Total.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "TOTAL_DEBIT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 2
		task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
	
	
If haveRecords("Personal Cash In Out Summary.IMD") And  haveRecords("Credit Total.IMD") And haveRecords("Debit Total.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Summary.IMD")
	Set task = db.VisualConnector
		id0 = task.AddDatabase ("Personal Cash In Out Summary.IMD")
		id1 = task.AddDatabase ("Credit Total.IMD")
		id2 = task.AddDatabase ("Debit Total.IMD")
		task.MasterDatabase = id0
		task.AppendDatabaseNames = FALSE
		task.IncludeAllPrimaryRecords = FALSE
		task.AddRelation id0, "ACCOUNT_NUMBER", id1, "ACCOUNT_NUMBER"
		task.AddRelation id0, "ACCOUNT_NUMBER", id2, "ACCOUNT_NUMBER"
		task.AddFieldToInclude id0, "ACCOUNT_NUMBER"
		task.AddFieldToInclude id0, "NO_OF_TRANSACTIONS"
		task.AddFieldToInclude id0, "TOTAL_TRANSACTION_AMOUNT"
		task.AddFieldToInclude id0, "ACCOUNT_NAME"
		task.AddFieldToInclude id0, "ACCOUNT_TYPE"
		task.AddFieldToInclude id0, "ROLE_TYPE"
		task.AddFieldToInclude id0, "ACCOUNT_STATUS"
		task.AddFieldToInclude id0, "CUSTOMER_NUMBER"
		task.AddFieldToInclude id0, "CUSTOMER_NAME"
		task.AddFieldToInclude id0, "CUSTOMER_TYPE"
		task.AddFieldToInclude id0, "TRANSACTION_DATE"
		task.AddFieldToInclude id0, "RUN_DATE"
		task.AddFieldToInclude id0, "TRANSACTION_CURRENCY"
		task.AddFieldToInclude id1, "TOTAL_CREDIT"
		task.AddFieldToInclude id2, "TOTAL_DEBIT"
		task.CreateVirtualDatabase = False
		dbName = "Personal Cash In Out Summary Final.IMD"
		task.OutputDatabaseName = dbName
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
End  Function

Function K_Export
If haveRecords("Personal Cash In Out Summary Final.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Summary Final.IMD") 
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "ACCOUNT_TYPE"
		task.AddFieldToInc "ROLE_TYPE"
		task.AddFieldToInc "ACCOUNT_STATUS"
		task.AddFieldToInc "CUSTOMER_NUMBER"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "RUN_DATE"
		task.AddFieldToInc "NO_OF_TRANSACTIONS"
		task.AddFieldToInc "TOTAL_TRANSACTION_AMOUNT"
		task.AddFieldToInc "TRANSACTION_CURRENCY"
		task.AddFieldToInc "TOTAL_CREDIT"
		task.AddFieldToInc "TOTAL_DEBIT"
		eqn = ""
		task.PerformTask Client.WorkingDirectory &  "Reports\Personal Cash In Out Summary.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("Personal Cash In Out Details.IMD") Then
	Set db = Client.OpenDatabase("Personal Cash In Out Details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "CUSTOMER_NUMBER"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "TRANSACTION_AMOUNT"
		task.AddFieldToInc "ORIGINAL_TRANSACTION_AMOUNT"
		task.AddFieldToInc "ORIGINAL_CURRENCY"
		task.AddFieldToInc "DIRECTION_OF_TRANSACTION"
		task.AddFieldToInc "TRANSACTION_CODE"
		task.AddFieldToInc "DESCRIPTION"
		task.AddFieldToInc "REFERENCE"
		task.AddFieldToInc "RUN_DATE"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\Personal Cash In Out Details.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Z_CleanUp
	DeleteFile ("Personal Cash In Out.IMD")
	DeleteFile ("First Personal Cash In.IMD")
	DeleteFile ("Personal Cash InOut First CashIn.IMD")
	DeleteFile ("Personal Cash Out After In.IMD")
	DeleteFile ("Personal Cash In Out Details.IMD")
	DeleteFile ("Personal Cash In Out Summary.IMD")
	DeleteFile ("Personal Cash In Out Summ.IMD")
	DeleteFile ("Personal Cash In Out Temp.IMD")
	DeleteFile ("Personal Cash In Out Temp2.IMD")
	DeleteFile("Personal Cash In Out Summary Final.IMD") 
	DeleteFile ("Credit Total.IMD")
	DeleteFile ("Debit Total.IMD")
End Function

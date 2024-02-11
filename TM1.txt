'====================================================================================================
'	Test#: 		TM1
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	12/28/2014 3:57:06 PM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "TM1"
Const scriptname_log ="TM1.iss"
Global errors_string As String
Const division = "UTC"
Dim fso As Object

Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Client.CloseAll
	Call JCleanUp()
	
	Call ASummHist_Average()		'Summarize History to get Customer Average by Transaction Type
	Call CCreateSumm_Today()
	Call DExtResults_INT()			'Join to get File to apply criteria to get results.
	Call EExtResults()			' Extract exceptions based on criteria.
	Call GSumm_Dtls()
	Call HRename_ResultFields()
	Call IExportDatabase()
	Call JCleanUp()
	
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

'Summarize History to get Customer Average by Transaction Type
Function ASummHist_Average
If haveRecords("Tran_Hist_Average.IMD") Then
	Set db = Client.OpenDatabase("Tran_Hist_Average.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "TRANSACTION_TYPE"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.Criteria = "@Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE"" .AND. TRANSACTION_AMOUNT <> 0.00"
	dbName = "Summ_Hist_Average.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM + SM_AVERAGE
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

'Create Join Key in Today's Database
Function CCreateSumm_Today
If haveRecords("Daily_Transactions_Today.IMD") Then
Set db = Client.OpenDatabase("Daily_Transactions_Today.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "TRANSACTION_TYPE"
	task.AddFieldToSummarize "POST_DATE"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "BRANCH_NAME"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.Criteria = "UTCID <> """" .AND. @Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE"" .AND. TRANSACTION_AMOUNT <> 0.00"
	dbName = "Summ_Tran_Today.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Summ_Tran_Today.IMD") Then
Set db = Client.OpenDatabase("Summ_Tran_Today.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_SUMMARY"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function


' Join to get File to apply criteria to get results.
Function DExtResults_INT
If haveRecords("Summ_Tran_Today.IMD") Then
Set db = Client.OpenDatabase("Summ_Tran_Today.IMD")
	Set task = db.JoinDatabase
   If haveRecords("Summ_Hist_Average.IMD") Then
	task.FileToJoin "Summ_Hist_Average.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TRANSACTION_AMOUNT_AVERAGE"
	task.AddSFieldToInc "TRANSACTION_AMOUNT_SUM"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.AddMatchKey "TRANSACTION_TYPE", "TRANSACTION_TYPE", "A"
	task.CreateVirtualDatabase = False
	dbName = "TM1_INT.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function


' Extract exceptions based on criteria.
Function EExtResults
If haveRecords("TM1_INT.IMD") Then
Set db = Client.OpenDatabase("TM1_INT.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "TM1_SUMM.IMD"
'	task.AddExtraction dbName, "", "((TRANSACTION_TYPE == ""SALE"" .AND. TRANSACTION_SUMMARY > (TRANSACTION_AMOUNT_AVERAGE * " & e_TM1_SA_Threshold_Pct & " ))  .OR. ( TRANSACTION_TYPE == ""REPO"" .AND. TRANSACTION_SUMMARY > (TRANSACTION_AMOUNT_AVERAGE  *  " & e_TM1_RP_Threshold_Pct & " ))) .AND. TRANSACTION_SUMMARY > " & e_TM1_Min_Value & " .AND.  TRANSACTION_AMOUNT_AVERAGE > " & e_TM1_Min_Value
	task.AddExtraction dbName, "", "TRANSACTION_TYPE == ""SALE"" .AND. TRANSACTION_SUMMARY > (TRANSACTION_AMOUNT_AVERAGE * " & e_TM1_SA_Threshold_Pct & " ) .AND. TRANSACTION_SUMMARY > " & e_TM1_Min_Value & " .AND.  TRANSACTION_AMOUNT_AVERAGE > " & e_TM1_Min_Value
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Analysis: Summarization
Function GSumm_Dtls
If haveRecords("Daily_Transactions_Today.IMD") Then
Set db = Client.OpenDatabase("Daily_Transactions_Today.IMD")
	Set task = db.JoinDatabase
   If haveRecords("TM1_SUMM.IMD") Then
	task.FileToJoin "TM1_SUMM.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "TRANSACTION_AMOUNT_SUM"
	task.AddSFieldToInc "TRANSACTION_AMOUNT_AVERAGE"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.AddMatchKey "TRANSACTION_TYPE", "TRANSACTION_TYPE", "A"
	task.Criteria = "@Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE"" .AND. TRANSACTION_AMOUNT <> 0.00"
	task.CreateVirtualDatabase = False
	dbName = "TM1_DTLS.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Data: Index Database
Function HRename_ResultFields
If haveRecords("TM1_SUMM.IMD") Then
Set db = Client.OpenDatabase("TM1_SUMM.IMD")

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_TRANSACTIONS"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_AVERAGE"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_AVERAGE", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_SUMMARY"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_SUM", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing

If haveRecords("TM1_DTLS.IMD") Then
Set db = Client.OpenDatabase("TM1_DTLS.IMD")

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_AVERAGE"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_AVERAGE", field
	task.PerformTask
		
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

Function IExportDatabase
If haveRecords("TM1_SUMM.IMD") Then
Set db = Client.OpenDatabase("TM1_SUMM.IMD")
	Set task = db.ExportDatabase
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "POST_DATE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "NO_OF_TRANSACTIONS"
	task.AddFieldToInc "TRANSACTION_SUMMARY"
	task.AddFieldToInc "TRANSACTION_TYPE"
	task.AddFieldToInc "CUSTOMER_AVERAGE"
	task.AddFieldToInc "CUSTOMER_SUMMARY"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "BRANCH_NAME"	
	eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\TM1_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
End If
	Set task = Nothing

If haveRecords("TM1_DTLS.IMD") Then
Set db = Client.OpenDatabase("TM1_DTLS.IMD")
	Set task = db.ExportDatabase
	Set task = db.ExportDatabase
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "ACCT_NO"
	task.AddFieldToInc "TRUST_CODE"
	task.AddFieldToInc "PAYMENT_TYPE"
	task.AddFieldToInc "TRANSACTION_CHANNEL"
	task.AddFieldToInc "TRANSACTION_BRANCH"
	task.AddFieldToInc "BRANCH_DESCRIPTION"
	task.AddFieldToInc "TRANSACTION_TYPE"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "POST_DATE"
	task.AddFieldToInc "TRAN_DATE"
	task.AddFieldToInc "UNITS"
	task.AddFieldToInc "RECEIPT_NUMBER"
	task.AddFieldToInc "UNIT_PRICE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "CUSTOMER_NAME"
	task.AddFieldToInc "AGENT_CODE"
	task.AddFieldToInc "NARRATIVE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "ADDRESS1"
	task.AddFieldToInc "ADDRESS2"
	task.AddFieldToInc "ADDRESS3"
	task.AddFieldToInc "DATE_OF_BIRTH"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "BRANCH_NAME"
	eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\TM1_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing	
End If
	Set task = Nothing
End Function

Function JCleanUp
	DeleteFile("Summ_Hist_Average.IMD")
	DeleteFile("Summ_Tran_Today.IMD")
	DeleteFile("TM1_INT.IMD")
	DeleteFile("TM1_SUMM.IMD")
	DeleteFile("TM1_DTLS.IMD")
End Function

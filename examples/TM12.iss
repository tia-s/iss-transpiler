'====================================================================================================
'	Test#: 		TM12
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	5/17/2015 12:46:57 AM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "TM12"
Const scriptname_log ="TM12.iss"
Global errors_string As String
Const division = "UTC"
Dim fso As Object

Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

'	Call P_Cleanup()

	Call A_JoinCustTransactions()
	Call B_SummWeeklyTransDay()
	Call C_SummWeeklyTransExclude()
	Call F_AboveDollarThreshold()
	Call G_Rename_SummFields()
	Call H_SummaryDetails()	
	Call K_ExportDatabase()
	Call L_ExportSumm()

'	Call P_Cleanup()
	
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

Function A_JoinCustTransactions
 If haveRecords("History_Transaction_History.IMD") Then
	Set db = Client.OpenDatabase("History_Transaction_History.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Monthly_History12.IMD"
	task.AddExtraction dbName, "", "UTCID <> """" .AND. @Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE""  .AND. @Age(@Date(),TRAN_DATE)<31"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If

 If haveRecords("Monthly_History12.IMD") Then
 	Set db = Client.OpenDatabase("Monthly_History12.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Monthly History TM12.IMD"
'	task.AddExtraction dbName, "", " .NOT. @Match(@Left(RECEIPT_NUMBER,2), ""VR"",""VX"",""VG"",""VC"",""VS"",""VH"",""VT"",""VV"",""VB"",""VM"",""VN"",""VA"",""VP"",""VD"",""VJ"",""VM"")"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	
 	Set db = Client.OpenDatabase("Monthly History TM12.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RUN_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Date()"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
 
	Set field = Nothing
	
	
 End If
 
If haveRecords("Monthly History TM12.IMD") And haveRecords("REPORTSVR.VDPUCID01.IMD") Then
	Set db = Client.OpenDatabase("Monthly History TM12.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "REPORTSVR.VDPUCID01.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.Criteria = "UTCID <> """" .AND. TRANSACTION_TYPE == ""SALE"" .AND. PAYMENT_TYPE <> ""BALANCE"" .AND.  .NOT. @Isini( ""UTC ATM"", CUSTOMER_NAME)"
	task.CreateVirtualDatabase = False
	dbName = "Deposits.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
End If
	Set task = Nothing
	Set db = Nothing
 
End Function


Function B_SummWeeklyTransDay
If haveRecords("Deposits.IMD") Then
Set db = Client.OpenDatabase("Deposits.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "TRAN_DATE"
	task.AddFieldToSummarize "RUN_DATE"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "BRANCH_NAME"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "STAFFIND"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.Criteria = ""
	dbName = "TM12_SUMM_INT.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
End Function


Function C_SummWeeklyTransExclude
If haveRecords("TM12_SUMM_INT.IMD") Then
Set db = Client.OpenDatabase("TM12_SUMM_INT.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "UTCID"
	task.AddFieldToSummarize "RUN_DATE"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "BRANCH_NAME"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "STAFFIND"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_SUM"
	task.AddFieldToTotal "NO_OF_RECS"
	task.Criteria = "(TRANSACTION_AMOUNT_SUM <= " & e_TM12_Thresh & ")"
	dbName = "TM12_SUMM_INT2.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
End Function


Function F_AboveDollarThreshold
If haveRecords("TM12_SUMM_INT2.IMD") Then
	Set db = Client.OpenDatabase("TM12_SUMM_INT2.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "TM12_SUMM.IMD"
	task.AddExtraction dbName, "", "(TRANSACTION_AMOUNT_SUM_SUM > " & e_TM12_Thresh & ")"

	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
End Function


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
	field.Equation = "@if(@Alltrim(STAFFIND) == ""Y"" , ""Yes"", ""No"")"
	field.Length = 50
	task.AppendField field
	task.PerformTask
End If
End Function



Function H_SummaryDetails
If haveRecords("Monthly History TM12.IMD") Then
	Set db = Client.OpenDatabase("Monthly History TM12.IMD")
	Set task = db.JoinDatabase
   If haveRecords("TM12_SUMM.IMD") Then
	task.FileToJoin "TM12_SUMM.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.Criteria = "UTCID <> """" .AND.  TRANSACTION_TYPE == ""SALE"""
	task.CreateVirtualDatabase = False
	dbName = "TM12_DTLS.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
End If
	Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
End Function

Function K_ExportDatabase
If haveRecords("TM12_DTLS.IMD") Then
Set db = Client.OpenDatabase("TM12_DTLS.IMD")
	Set task = db.ExportDatabase
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "ACCT_NO"
	task.AddFieldToInc "TRUST_CODE"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "BRANCH_NAME"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "RECEIPT_NUMBER"
	task.AddFieldToInc "TRANSACTION_BRANCH"
	task.AddFieldToInc "BRANCH_DESCRIPTION"
	task.AddFieldToInc "TRANSACTION_TYPE"
	task.AddFieldToInc "TRAN_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "UNITS"
	task.AddFieldToInc "UNIT_PRICE"
	task.AddFieldToInc "POST_DATE"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "RISK_RATING"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RUN_DATE"	
	task.AddFieldToInc "PAYMENT_TYPE"		
eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\TM12_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
Set db = Nothing
End If
Set task = Nothing
End Function

Function L_ExportSumm	
If haveRecords("TM12_SUMM.IMD") Then
Set db = Client.OpenDatabase("TM12_SUMM.IMD")
	Set task = db.ExportDatabase
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "BRANCH_NAME"
	task.AddFieldToInc "STAFF"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "RATING_SOURCE"
	task.AddFieldToInc "RISK_RATING"
'	task.AddFieldToInc "POST_DATE"
	task.AddFieldToInc "RUN_DATE"
	task.AddFieldToInc "TRANSACTION_DAYS"
	task.AddFieldToInc "NO_OF_TRANSACTIONS"
	task.AddFieldToInc "TRANSACTION_SUMMARY"
	
	eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\TM12_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
End If
	Set task = Nothing
End Function


Function P_Cleanup
	DeleteFile("Deposits.IMD")
	DeleteFile("Monthly History TM12.IMD")
	DeleteFile("Monthly History12.IMD")
	DeleteFile("TM12_SUMM_INT.IMD")
	DeleteFile("TM12_SUMM_INT2.IMD")
	DeleteFile("TM12_SUMM.IMD")
	DeleteFile("TM12_DTLS.IMD")
End Function


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





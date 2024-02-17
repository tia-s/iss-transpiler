'====================================================================================================
'	Test#: 		AML13 - Withdrawals from a previously dormant Account
'	Risk:		Activity on previously dormant /inactive account may be suspicious.
' 	Objective:	Identify large cash withdrawals from a previously dormant/inactive account.
'			Note: Watch reactivated dormant account For a period of Time.

' 	Frequency:	daily
' 	Last Modified:	24/03/2014 11:16:00 AM
'====================================================================================================
'	Script Dependencies: Import, Interim
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "AML13 - Withdrawals from a previously dormant Account"
Const scriptname_log ="AML13 - Withdrawals from a previously dormant Account.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main

	Ignorewarning(True)
	Client.CloseAll
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Get_Withdrawals_from_Previously_Dormant_Acc
	Client.CloseAll
	Get_Withdrawal_from_Dormant_Over_Threshold
	Client.CloseAll
	Get_Withdrawal_Dormant_Acc_Summary
	Client.CloseAll
	Get_Withdrawal_Dormant_Acc_Detail
	Client.CloseAll
	Join_RiskRating
	Client.CloseAll
	ExportDatabase	
	
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


Function Get_Withdrawals_from_Previously_Dormant_Acc
If haveRecords("General Acct Master Lite.IMD") Then
	Set db = Client.OpenDatabase("General Acct Master Lite.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "FORACID"
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "ACCT_NAME"
			task.AddFieldToInc "CUST_NAME"
			task.AddFieldToInc "SCHM_DESC"
			task.AddFieldToInc "CUSTOMER_TYPE"
			dbName = "GAM Extract.IMD"
			task.AddExtraction dbName, "", ""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
If haveRecords("Dormant Accounts Activated Large.IMD") Then
	Set db = Client.OpenDatabase("Dormant Accounts Activated Large.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "AUDIT_DATE_DATE"
			task.AddFieldToInc "AUDIT_DATE_TIME"
			task.AddFieldToInc "COMP_STATUS_CHANGE"
			task.AddFieldToInc "COMP_CHANGE_FROM"
			task.AddFieldToInc "COMP_CHANGE_TO"
			dbName = "Dormant Account Activated Extract.IMD"
			task.AddExtraction dbName, "", ""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
If haveRecords("Customer Induced Txns.IMD") And haveRecords("Dormant Account Activated Extract.IMD") Then
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "Dormant Account Activated Extract.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "AUDIT_DATE_DATE"
			task.AddSFieldToInc "AUDIT_DATE_TIME"
			task.AddSFieldToInc "COMP_STATUS_CHANGE"
			task.AddSFieldToInc "COMP_CHANGE_FROM"
			task.AddSFieldToInc "COMP_CHANGE_TO"
			task.AddMatchKey "ACID", "ACID", "A"
			'task.Criteria = "PART_TRAN_TYPE ==""D"" .AND.INSTRMNT_TYPE <> ""CHQ""  .AND.  TRAN_TYPE == ""C"""
			task.Criteria = ""
			task.CreateVirtualDatabase = False
			dbName = "Withdraw_from_Previous_Dormat_Account_Int.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
If haveRecords("Withdraw_from_Previous_Dormat_Account_Int.IMD") And haveRecords("GAM Extract.IMD") Then
	Set db = Client.OpenDatabase("Withdraw_from_Previous_Dormat_Account_Int.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "GAM Extract.IMD"
		task.IncludeAllPFields
		task.IncludeAllSFields
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Withdrawal_from_Previously_Dormat_Account.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function Get_Withdrawal_from_Dormant_Over_Threshold

If haveRecords("Withdrawal_from_Previously_Dormat_Account.IMD") Then
	Set db = Client.OpenDatabase("Withdrawal_from_Previously_Dormat_Account.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "FORACID"
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "ACCT_NAME"
			task.AddFieldToInc "CUST_NAME"
			task.AddFieldToInc "SCHM_DESC"
			task.AddFieldToInc "CUSTOMER_TYPE"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "TRAN_ID"
			task.AddFieldToInc "BR_CODE"
			task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddFieldToInc "SOL_ID"
			task.AddFieldToInc "FX_RATE_TO_JMD"
			task.AddFieldToInc "TRANSACTION_AMT"
			task.AddFieldToInc "TRANSACTION_AMOUNT_JM"
			task.AddFieldToInc "FX_RATE_JMD_TO_USD"
			task.AddFieldToInc "TRANSACTION_AMOUNT_US"
			task.AddFieldToInc "BR_NAME"
			task.AddFieldToInc "AUDIT_DATE_DATE"
			task.AddFieldToInc "AUDIT_DATE_TIME"
			task.AddFieldToInc "COMP_STATUS_CHANGE"
			task.AddFieldToInc "COMP_CHANGE_FROM"
			task.AddFieldToInc "COMP_CHANGE_TO"
			dbName = "Withdrawal_Prev_Dormant_Acc.IMD"
			'task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US >= " & e_AML13_Trans_Value_US
			task.AddExtraction dbName, "", ""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If	
End Function

Function Get_Withdrawal_Dormant_Acc_Summary
  If haveRecords("Withdrawal_Prev_Dormant_Acc.IMD") Then
	Set db = Client.OpenDatabase("Withdrawal_Prev_Dormant_Acc.IMD")
		Set task = db.Summarization
			task.AddFieldToSummarize "FORACID"
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "ACCT_NAME"
			task.AddFieldToInc "CUST_NAME"
			task.AddFieldToInc "AUDIT_DATE_DATE"
			task.AddFieldToInc "AUDIT_DATE_TIME"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "COMP_CHANGE_FROM"
			task.AddFieldToInc "COMP_CHANGE_TO"
			task.AddFieldToTotal "TRANSACTION_AMOUNT_JM"
			task.AddFieldToTotal "TRANSACTION_AMOUNT_US"
			dbName = "AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD"
			task.OutputDBName = dbName
			task.CreatePercentField = FALSE
			task.UseFieldFromFirstOccurrence = TRUE
			task.StatisticsToInclude = SM_SUM
			task.PerformTask
		Set task = Nothing
	Set db = Nothing
   End If

     If haveRecords("AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD") Then
	Set db = Client.OpenDatabase("AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "ACCT_NAME", field

	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "CUST_NAME", field
	
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
		field.Name = "NO_OF_TRANSACTIONS"
		field.Description = "Number of records found for this key value"
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 0
		task.ReplaceField "NO_OF_RECS", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_AMOUNT_JM"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 2
		task.ReplaceField "TRANSACTION_AMOUNT_JM_SUM", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_AMOUNT_US"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 2
		task.ReplaceField "TRANSACTION_AMOUNT_US_SUM", field

	Set field = db.TableDef.NewField
		field.Name = "AUDIT_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "AUDIT_DATE_DATE", field

	Set field = db.TableDef.NewField
		field.Name = "CHANGE_FROM"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 50
		task.ReplaceField "COMP_CHANGE_FROM", field

	Set field = db.TableDef.NewField
		field.Name = "CHANGE_TO"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 50
		task.ReplaceField "COMP_CHANGE_TO", field
		
	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "FORACID", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
  End If
End Function

Function Get_Withdrawal_Dormant_Acc_Detail
     If haveRecords("Withdrawal_Prev_Dormant_Acc.IMD") Then
	Set db = Client.OpenDatabase("Withdrawal_Prev_Dormant_Acc.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "FORACID"
			task.AddFieldToInc "SOL_ID"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "TRAN_ID"
			task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddFieldToInc "TRANSACTION_AMT"
			task.AddFieldToInc "FX_RATE_TO_JMD"
			task.AddFieldToInc "TRANSACTION_AMOUNT_JM"
			task.AddFieldToInc "FX_RATE_JMD_TO_USD"
			task.AddFieldToInc "TRANSACTION_AMOUNT_US"
			task.AddFieldToInc "BR_NAME"
			dbName = "AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD"
			task.AddExtraction dbName, "", ""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If

     If haveRecords("AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD") Then
	Set db = Client.OpenDatabase("AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "TRAN_ID", field
	
	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_CURRENCY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "TRANSACTION_CRNCY_CODE", field

	Set field = db.TableDef.NewField
		field.Name = "BRANCH_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "BR_NAME", field
		
	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "FORACID", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
   End If
End Function



Function Join_RiskRating
If haveRecords("Account_Turnover_wRisk.IMD") And haveRecords("AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD") Then
	Set db = Client.OpenDatabase("AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD")	
	Set task = db.JoinDatabase
		task.FileToJoin "Account_Turnover_wRisk.IMD" 
		task.IncludeAllPFields
		task.AddSFieldToInc "MONTHLY_DEPOSIT"
		task.AddSFieldToInc "RISK_SCORE"
		task.AddSFieldToInc "RISK_CATEGORY"
		task.AddSFieldToInc "OCCUPATION"
		task.AddSFieldToInc "INDUSTRY"
		task.AddSFieldToInc "ANNUAL_INCOME"
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "AML13_Withdrawal_Prev_Dormant_Account.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExportDatabase
If haveRecords("AML13_Withdrawal_Prev_Dormant_Account.IMD") Then
	Set db = Client.OpenDatabase("AML13_Withdrawal_Prev_Dormant_Account.IMD")
	Set task = db.ExportDatabase
		task.IncludeAllFields
		eqn = ""
		task.PerformTask Client.Workingdirectory & "Reports\AML13_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
	RESULTSLOG(db.name)
	Set db = Nothing
	Set task = Nothing
	Else 
NORESULTSLOG("AML13_SUMMARY.IMD") 
End If

If haveRecords("AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD") Then
	Set db = Client.OpenDatabase("AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD")
		Set task = db.ExportDatabase
		task.IncludeAllFields
			eqn = ""
			task.PerformTask Client.Workingdirectory & "Reports\AML13_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
			Set task = Nothing
	Set db = Nothing
End If
End Function

Function CleanUp
	DeleteFile("GAM Extract.IMD") 
	DeleteFile("AML13_Withdrawal_Prev_Dormant_Account_Detail.IMD") 
	DeleteFile("Withdraw_from_Previous_Dormat_Account.IMD") 
	DeleteFile("Withdraw_from_Previous_Dormat_Account_Int.IMD") 
	DeleteFile("Withdrawal_Prev_Dormant_Acc.IMD")
	'DeleteFile("AML13_Withdrawal_Prev_Dormant_Account.IMD")
	DeleteFile("AML13_Withdrawal_Prev_Dormant_Account_tmp.IMD")
	DeleteFile("Dormant Account Activated Extract.IMD")
	DeleteFile("Withdrawal_from_Previously_Dormat_Account.IMD")
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
LogName = "Results.csv"

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
LogName = "Results.csv"

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



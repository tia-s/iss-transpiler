'====================================================================================================
'	Test#: 		AML15 - Large Early Loan Payments
'	Risk:		Sudden loan payments may be suspicious.
' 	Objective:	Identify customers who repay delinquent loans unexpectedly.
' 	Frequency:	Weekly
' 	Last Modified:	13-Nov-14
'	Comments:	Change to Consolidate loan payments for comparison with the threshold.
'====================================================================================================
'	Script Dependencies: Import, interim
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "AML15 - Large Early Loan Payments"
Const scriptname_log ="AML15 - Large Early Loan Payments.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object


Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine
	
	Client.CloseAll
	CleanUp
	
	Get_Active_Loan_Not_Past_Due
	JoinwklyTxn_DelLoans
	JoinActive_To_GAM
	Client.CloseAll
	ModifyFieldForExport
	ActiveLoanTxnGAM_Sum
	ExtractLargeLoanPayments
	Join_RiskRating
 	
 	Client.CloseAll
	ExportResultSet
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


Function Get_Active_Loan_Not_Past_Due
If haveRecords("TBAADM.GEN_ACCT_CLASS_TABLE.IMD") And haveRecords("Active Loans.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.GEN_ACCT_CLASS_TABLE.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Active Loans.IMD"
		task.AddPFieldToInc "ACID"
		task.AddSFieldToInc "REP_SHDL_DATE_DATE"
		task.AddSFieldToInc "DIS_AMT"
		task.AddSFieldToInc "REP_PERD_MTHS"
		task.AddSFieldToInc "REPHASEMENT_PRINCIPAL"
		task.AddMatchKey "ACID", "ACID", "A"
		task.Criteria = "PD_FLG <> ""Y"""
		task.CreateVirtualDatabase = False
		dbName = "Active Loan Not PD.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM_GEN_ACCT_CLASS_TABLE.IMD or Active Loans.IMD", "Get_Active_Loan_Not_Past_Due", "JoinDatabase", "Error", "Databases empty or does not exist.")

End If

End Function

' File: Join Databases
Function JoinwklyTxn_DelLoans
If haveRecords("Customer Induced Txns.IMD") And haveRecords("Active Loan Not PD.IMD") Then
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Active Loan Not PD.IMD"
		task.IncludeAllPFields
		task.IncludeAllSFields
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "ActiveLoanTransactions.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Customer Induced Txns.IMD or Active Loan Not PD.IMD", "JoinwklyTxn_DelLoans", "JoinDatabase", "Error", "Databases empty or does not exist.")

End If
End Function

Function JoinActive_To_GAM
If haveRecords("ActiveLoanTransactions.IMD") And haverecords("General Acct Master Lite.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTransactions.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "General Acct Master Lite.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "FORACID"
		task.AddSFieldToInc "SCHM_DESC"
		task.AddSFieldToInc "ACCT_NAME"
		task.AddSFieldToInc "CUSTOMER_TYPE"
		task.AddSFieldToInc "CLR_BAL_AMT"
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "ActiveLoanTxnGAM Int.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("ActiveLoanTransactions.IMD or General Acct Master Lite.IMD", "JoinActive_To_GAM", "JoinDatabase", "Error", "Databases empty or does not exist.")

End If
		
If haveRecords("ActiveLoanTxnGAM Int.IMD") And haveRecords("NCB_BRANCHES.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTxnGAM Int.IMD")
	Set task = db.JoinDatabase
	   	task.FileToJoin "NCB_BRANCHES.IMD"
	   	task.IncludeAllPFields
	   	task.AddSFieldToInc "BR_NAME"
	   	task.AddMatchKey "SOL_ID", "MICR_BRANCH_CODE", "A"
		task.CreateVirtualDatabase = False
		dbName = "ActiveLoanTxnGAM.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_PRIM
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("ActiveLoanTxnGAM Int.IMD or NCB_BRANCHES.IMD", "JoinActive_To_GAM", "JoinDatabase", "Error", "Databases empty or does not exist.")

End If
End Function

' Modify Field
Function ModifyFieldForExport
If haveRecords("ActiveLoanTxnGAM.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTxnGAM.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "TRAN_ID", field

	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_AMOUNT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_PARTICULARS"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 50
	task.ReplaceField "TRAN_PARTICULAR", field

	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_CURRECNY"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "TRANSACTION_CRNCY_CODE", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_NUMBER"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "FORACID", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "ACCT_NAME", field
	
	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_BRANCH"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 30
	task.ReplaceField "BR_NAME", field
	
	Set field = db.TableDef.NewField
	field.Name = "DISBURSED_AMOUNT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "DIS_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "REPAYMENT_SCHEDULE_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "REP_SHDL_DATE_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "REPAYMENT_PERIOD_IN_MONTHS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "REP_PERD_MTHS", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_DESCRIPTION"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "SCHM_DESC", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_BALANCE"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "CLR_BAL_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "PRINCIPAL_AMOUNT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "REPHASEMENT_PRINCIPAL", field
	task.PerformTask
	
	Set field = db.TableDef.NewField
	field.Name = "RUNDATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Date()"
	task.AppendField field
	task.PerformTask
	
	
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

End If
End Function




' Analysis: Summarization to total loan payments for period
Function ActiveLoanTxnGAM_Sum
If haveRecords("ActiveLoanTxnGAM.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTxnGAM.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "ACCOUNT_NUMBER"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRAN_DATE_TIME"
	task.AddFieldToInc "TRANSACTION_ID"
	task.AddFieldToInc "PART_TRAN_SRL_NUM"
	task.AddFieldToInc "DEL_FLG"
	task.AddFieldToInc "TRAN_TYPE"
	task.AddFieldToInc "TRAN_SUB_TYPE"
	task.AddFieldToInc "PART_TRAN_TYPE"
	task.AddFieldToInc "GL_SUB_HEAD_CODE"
	task.AddFieldToInc "ACID"
	task.AddFieldToInc "TRANSACTION_PARTICULARS"
	task.AddFieldToInc "RPT_CODE"
	task.AddFieldToInc "REF_NUM"
	task.AddFieldToInc "INSTRMNT_TYPE"
	task.AddFieldToInc "INSTRMNT_NUM"
	task.AddFieldToInc "INSTRMNT_ALPHA"
	task.AddFieldToInc "TRAN_RMKS"
	task.AddFieldToInc "PSTD_FLG"
	task.AddFieldToInc "PRNT_ADVC_IND"
	task.AddFieldToInc "AMT_RESERVATION_IND"
	task.AddFieldToInc "RESERVATION_AMT"
	task.AddFieldToInc "RESTRICT_MODIFY_IND"
	task.AddFieldToInc "RCRE_TIME_TIME"
	task.AddFieldToInc "CUSTOMER_ID"
	task.AddFieldToInc "VOUCHER_PRINT_FLG"
	task.AddFieldToInc "MODULE_ID"
	task.AddFieldToInc "BR_CODE"
	task.AddFieldToInc "FX_TRAN_AMT"
	task.AddFieldToInc "RATE_CODE"
	task.AddFieldToInc "RATE"
	task.AddFieldToInc "CRNCY_CODE"
	task.AddFieldToInc "NAVIGATION_FLG"
	task.AddFieldToInc "REF_CRNCY_CODE"
	task.AddFieldToInc "REF_AMT"
	task.AddFieldToInc "SOL_ID"
	task.AddFieldToInc "BANK_CODE"
	task.AddFieldToInc "TREA_REF_NUM"
	task.AddFieldToInc "TREA_RATE"
	task.AddFieldToInc "TS_CNT"
	task.AddFieldToInc "TRAN_PARTICULAR_2"
	task.AddFieldToInc "TRAN_PARTICULAR_CODE"
	task.AddFieldToInc "REVERSAL_FLG"
	task.AddFieldToInc "FX_RATE_TO_JMD"
	task.AddFieldToInc "USD"
	task.AddFieldToInc "ACCOUNT_BRANCH"
	task.AddFieldToInc "ACID1"
	task.AddFieldToInc "REPAYMENT_SCHEDULE_DATE"
	task.AddFieldToInc "DISBURSED_AMOUNT"
	task.AddFieldToInc "REPAYMENT_PERIOD_IN_MONTHS"
	task.AddFieldToInc "PRINCIPAL_AMOUNT"
	task.AddFieldToInc "ACCOUNT_DESCRIPTION"
	task.AddFieldToInc "ACCOUNT_NAME"
	task.AddFieldToInc "CUSTOMER_TYPE"
	task.AddFieldToInc "ACCOUNT_BALANCE"
	task.AddFieldToInc "BR_NAME1"
	task.AddFieldToInc "RUNDATE"
	task.AddFieldToTotal "TRANSACTION_AMOUNT"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_JM"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_US"
	task.Criteria = "TRAN_SUB_TYPE <> ""IC"" .AND. ACCOUNT_BALANCE <> 0 .AND. PART_TRAN_TYPE == ""C"" .AND. (.NOT. @Isini(""Loan Coll. From"", TRANSACTION_PARTICULARS))"
	dbName = "ActiveLoanTxnGAM_Sum.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
 end if
End Function



' Data: Direct Extraction
Function ExtractLargeLoanPayments
If haveRecords("ActiveLoanTxnGAM_Sum.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTxnGAM_Sum.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "Large_Loan_Repayments_Made_Early_tmp.IMD"
		'task.AddExtraction dbName, "", "( TRANSACTION_AMOUNT_JM_Sum >=" & e_AML15_TXN_VALUE & ") .AND. TRANSACTION_AMOUNT_Sum >= ( @Abs(ACCOUNT_BALANCE) *(" & e_Loan_Payment_Percent_Score & "/100)) "
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing

Else
	Call logfile("ActiveLoanTxnGAM.IMD", "ExtractLargeLoanPayments", "Direct Extraction", "Error", "Databases empty or does not exist.")

End If	

If haveRecords("Large_Loan_Repayments_Made_Early_tmp.IMD") Then
	Set db = Client.OpenDatabase("ActiveLoanTxnGAM.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Large_Loan_Repayments_Made_Early_tmp.IMD"
	task.IncludeAllPFields
	task.AddMatchKey "ACCOUNT_NUMBER", "ACCOUNT_NUMBER", "A"
	task.CreateVirtualDatabase = False
	dbName = "Large Loan payment details.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Large_Loan_Repayments_Made_Early_tmp.IMD") Then
	Set db = Client.OpenDatabase("Large_Loan_Repayments_Made_Early_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_PAYMENTS"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If	
End Function

Function Join_RiskRating
If haveRecords("Account_Turnover6.IMD") And haveRecords("Large_Loan_Repayments_Made_Early_tmp.IMD") Then
	Set db = Client.OpenDatabase("Large_Loan_Repayments_Made_Early_tmp.IMD")	
	Set task = db.JoinDatabase
		task.FileToJoin "Account_Turnover6.IMD" 
		task.IncludeAllPFields
		task.AddSFieldToInc "MONTHLY_DEPOSIT"
		task.AddSFieldToInc "RISK_SCORE"
		task.AddSFieldToInc "RISK_CATEGORY"
		task.AddSFieldToInc "OCCUPATION"
		task.AddSFieldToInc "INDUSTRY"
		task.AddSFieldToInc "ANTICIPATED_TURNOVER"
		task.AddMatchKey "CUSTOMER_ID", "CUST_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Large_Loan_Repayments_Made_Early.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

' File - Export Database
Function ExportResultSet

If haveRecords("Large_Loan_Repayments_Made_Early.IMD") Then
	Set db = Client.OpenDatabase("Large_Loan_Repayments_Made_Early.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "NO_OF_PAYMENTS"
		task.AddFieldToInc "TRANSACTION_AMOUNT_SUM"
		task.AddFieldToInc "TRANSACTION_AMOUNT_JM_SUM"
		task.AddFieldToInc "TRANSACTION_AMOUNT_US_SUM"
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "RATE"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "ACCOUNT_BRANCH"
		task.AddFieldToInc "REPAYMENT_SCHEDULE_DATE"
		task.AddFieldToInc "DISBURSED_AMOUNT"
		task.AddFieldToInc "REPAYMENT_PERIOD_IN_MONTHS"
		task.AddFieldToInc "PRINCIPAL_AMOUNT"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "ACCOUNT_BALANCE"
		task.AddFieldToInc "RUNDATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANTICIPATED_TURNOVER"
	eqn = ""
	task.PerformTask "Reports\AML15_SUM.MDB", "Database", "MDB2000", 1, db.Count, eqn
		RESULTSLOG(db.name)
	Set task = Nothing
	Set db = Nothing

Else 
NORESULTSLOG("AML15_SUMMARY.IMD") 
End If

If haveRecords("Large Loan payment details.IMD") Then
	Set db = Client.OpenDatabase("Large Loan payment details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "TRANSACTION_ID"
		task.AddFieldToInc "TRANSACTION_AMOUNT"
		task.AddFieldToInc "TRANSACTION_PARTICULARS"
		task.AddFieldToInc "TRANSACTION_CURRECNY"
		task.AddFieldToInc "TRANSACTION_AMOUNT_JM"
		task.AddFieldToInc "DISBURSED_AMOUNT"
		task.AddFieldToInc "REPAYMENT_SCHEDULE_DATE"
		task.AddFieldToInc "REPAYMENT_PERIOD_IN_MONTHS"
		task.AddFieldToInc "PRINCIPAL_AMOUNT"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "ACCOUNT_BRANCH"
		task.AddFieldToInc "ACCOUNT_BALANCE"
		task.AddFieldToInc "RUNDATE"
		eqn = ""
		task.PerformTask "Reports\AML15_Details.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function CleanUp
	DeleteFile("ActiveLoanTxnGAM.IMD")
	DeleteFile("ActiveLoanTransactions.IMD")
	'DeleteFile("Large_Loan_Repayments_Made_Early.IMD") 
	DeleteFile("Large_Loan_Repayments_Made_Early_tmp.IMD") 
	DeleteFile("ActiveLoanTxnGAM Int.IMD")
	DeleteFile("Active Loan Not PD.IMD")
	
	DeleteFile("ActiveLoanTxnGAM_Sum.IMD")
	DeleteFile("Large Loan payment details.IMD")
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

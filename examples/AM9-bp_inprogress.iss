'====================================================================================================
'	Test#: 		AM9 - New Joint Party with Significant Withdrawal
'	Risk:		Joint parties may be added to accounts followed by significant withdrawals.
' 	Objective:	Identify addition of joint parties to accounts followed by significant withdrawal activity.
' 	Frequency:	Daily
' 	Last Modified:	11-Nov-14 
'====================================================================================================
'	Script Dependencies: Import, interim
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "AM9 - New Joint Party with Significant Withdrawal"
Const scriptname_log ="AM9 - New Joint Party with Significant Withdrawal.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main

	Ignorewarning(True)
	Client.CloseAll
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine
	Client.CloseAll
	CleanUp
		
	ExtractModifiedAAS
	CompAcctSrlNum
	SortAASRecordAdded
	SummarizeAccounts
	ExtractJointPartiesAddedToday
	GetJointPartyMain
	JoinCustTxnNewJoint
	Get_Joint_Holder
	Get_Main_Account_Details
	Join_RiskRating
	RenameFieldsJoint
	RenameFieldsDetails
	RenameFieldsSummary
	ExportResults
	
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
' ANALYSIS
'---------------------------------------------------------------------------------------------------------------------------------------
' Get AAS records added (joint parties added)
Function ExtractModifiedAAS
If haveRecords("TBAADM.AUDIT_TABLE_LARGE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.AUDIT_TABLE_LARGE.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "AAS_Records_Added.IMD"
		task.AddExtraction dbName, "", "TABLE_NAME == ""AAS"" .AND. MODIFIED_FIELDS_DATA == ""RECORD ADDED|||"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Split table key to isolate the join party position number.
Function CompAcctSrlNum
If haveRecords("AAS_Records_Added.IMD") Then
	Set db = Client.OpenDatabase("AAS_Records_Added.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_ACCT_SRL_NUM"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@AllTrim(@Split(TABLE_KEY, """", ""/"", 2))"
		field.Length = 3
		task.AppendField field
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Sort and summarize to identify only instances where the joint party was added days after the main account holder
' ie, the joint was added after the account open date.
Function SortAASRecordAdded
If haveRecords("AAS_Records_Added.IMD") Then
	Set db = Client.OpenDatabase("AAS_Records_Added.IMD")
	Set task = db.Sort
		task.AddKey "ACID", "A"
		task.AddKey "COMP_ACCT_SRL_NUM", "A"
		task.AddKey "AUDIT_DATE_DATE", "D"
		dbName = "Sorted AAS Records Added.IMD"
		task.PerformTask dbName
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function SummarizeAccounts
If haveRecords("Sorted AAS Records Added.IMD") Then
	Set db = Client.OpenDatabase("Sorted AAS Records Added.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "ACID"
		task.IncludeAllFields
		dbName = "Summ AAS Records Added.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExtractJointPartiesAddedToday
If haveRecords("Summ AAS Records Added.IMD") Then
	Set db = Client.OpenDatabase("Summ AAS Records Added.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "ACID"
		dbName = "Joint_Parties_Acid.IMD"
		task.AddExtraction dbName, "", "COMP_ACCT_SRL_NUM > ""001""" 
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If

' collect all joint holders added for the accounts that do not have a main holder added
If haveRecords("Joint_Parties_Acid.IMD") And haveRecords("AAS_Records_Added.IMD") Then
	Set db = Client.OpenDatabase("AAS_Records_Added.IMD")	
	Set task = db.JoinDatabase
		task.FileToJoin "Joint_Parties_Acid.IMD" 
		task.AddPFieldToInc "AUDIT_DATE_DATE"
		task.AddPFieldToInc "AUDIT_DATE_TIME"
		task.AddPFieldToInc "ENTERER_ID"
		task.AddPFieldToInc "AUTH_ID"
		task.AddPFieldToInc "RMKS"
		task.AddPFieldToInc "COMP_ACCT_SRL_NUM"
		task.IncludeAllSFields
		task.AddMatchKey "ACID", "ACID", "A"
		task.Criteria = "COMP_ACCT_SRL_NUM > ""001"""
		task.CreateVirtualDatabase = False
		dbName = "Joint_Parties_Added_Today.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Joint_Parties_Added_Today.IMD") Then
	Set db = Client.OpenDatabase("Joint_Parties_Added_Today.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "ACCT_HOLDER_KEY"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@AllTrim( ACID)  +  ""-"" + @AllTrim(COMP_ACCT_SRL_NUM)"
		field.Length = 20
		task.AppendField field
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function GetJointPartyMain
' Collect joint holder information from ACCT_AUTH_SIGN_TABLE.	
If haveRecords("Joint_Parties_Added_Today.IMD") And haveRecords("Other Accounts - CGM Join.IMD") Then
	Set db = Client.OpenDatabase("Joint_Parties_Added_Today.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Other Accounts - CGM Join.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "CUST_NAME"
		task.AddSFieldToInc "CUST_ID"
		task.AddMatchKey "ACCT_HOLDER_KEY", "OTHER_ACCT_KEY", "A"
		task.CreateVirtualDatabase = False
		dbName = "Joint Parties Today Complete.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Joint Parties Today Complete.IMD") Then
	Set db = Client.OpenDatabase("Joint Parties Today Complete.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "JOINT_ACCOUNT_HOLDER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "CUST_NAME", field
	Set field = db.TableDef.NewField
		field.Name = "JOINT_CUSTOMER_ID"
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
	
If haveRecords("Joint Parties Today Complete.IMD") And haveRecords("General Acct Master Lite.IMD")  Then
	Set db = Client.OpenDatabase("Joint Parties Today Complete.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "General Acct Master Lite.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "FORACID"
		task.AddSFieldToInc "ACCT_NAME"
		task.AddSFieldToInc "CUST_ID"
		task.AddSFieldToInc "CUST_NAME"
		task.AddSFieldToInc "SCHM_TYPE"
		task.AddSFieldToInc "SCHM_DESC"
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Joint Parties Today Complete - GAM.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
End Function


' Get transactions on accounts with joint party added.
Function JoinCustTxnNewJoint
If haveRecords("Customer Induced Txns.IMD") And haveRecords("Joint Parties Today Complete - GAM.IMD") Then
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Joint Parties Today Complete - GAM.IMD"
		task.AddPFieldToInc "ACID"
		task.AddPFieldToInc "TRAN_ID"
		task.AddPFieldToInc "TRAN_TYPE"
		task.AddPFieldToInc "TRANSACTION_DATE"
		task.AddPFieldToInc "TRAN_SUB_TYPE"
		task.AddPFieldToInc "TRANSACTION_AMT"
		task.AddPFieldToInc "TRAN_PARTICULAR"
		task.AddPFieldToInc "TRAN_RMKS"
		task.AddPFieldToInc "SOL_ID"
		task.AddPFieldToInc "CUST_ID"
		task.AddPFieldToInc "TRANSACTION_CRNCY_CODE"
		task.AddSFieldToInc "FORACID"
		task.AddSFieldToInc "ACCT_NAME"
		task.AddSFieldToInc "CUST_ID"
		task.AddSFieldToInc "CUST_NAME"
		task.AddSFieldToInc "SCHM_TYPE"
		task.AddSFieldToInc "SCHM_DESC"
		task.AddMatchKey "ACID", "ACID", "A"
		'task.Criteria = "PART_TRAN_TYPE == ""D"" .AND. TRANSACTION_AMOUNT_JM >=  " & e_NewParty_Withdrawal_Value_JM
		task.Criteria = ""
		task.CreateVirtualDatabase = False
		dbName = "AM9_Wdrl_After_New_Joint_Dtls.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Join Joint Parties Today Complete to transaction
Function Get_Joint_Holder
If haveRecords("Joint Parties Today Complete.IMD") And haveRecords("AM9_Wdrl_After_New_Joint_Dtls.IMD") Then
	Set db = Client.OpenDatabase("Joint Parties Today Complete.IMD")	
	Set task = db.JoinDatabase
		task.FileToJoin "AM9_Wdrl_After_New_Joint_Dtls.IMD" 
		task.IncludeAllPFields
		task.AddSFieldToInc "FORACID"
		task.AddSFieldToInc "TRANSACTION_DATE"
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "AM9_Joint_Parties_tmp.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Summarize results and extract accounting information
Function Get_Main_Account_Details
If haveRecords("AM9_Wdrl_After_New_Joint_Dtls.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Dtls.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "FORACID"
		task.AddFieldToInc "ACCT_NAME"
		task.AddFieldToInc "CUST_ID"
		task.AddFieldToInc "ACID"
		task.AddFieldToInc "CUST_NAME"
		task.AddFieldToInc "SCHM_TYPE"
		task.AddFieldToInc "SCHM_DESC"
		task.AddFieldToInc "TRANSACTION_DATE"
		dbName = "AM9_Wdrl_After_New_Joint_Summ_tmp.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Join_RiskRating
If haveRecords("Account_Turnover_wRisk.IMD") And haveRecords("AM9_Joint_Parties_tmp.IMD") Then
	Set db = Client.OpenDatabase("AM9_Joint_Parties_tmp.IMD")	
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
		dbName = "AM9_Joint_Parties.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Account_Turnover_wRisk.IMD") And haveRecords("AM9_Wdrl_After_New_Joint_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Summ_tmp.IMD")	
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
		dbName = "AM9_Wdrl_After_New_Joint_Summ.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Rename fields to export
Function RenameFieldsJoint
If haveRecords("AM9_Joint_Parties.IMD") Then
	Set db = Client.OpenDatabase("AM9_Joint_Parties.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "AUDIT_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "AUDIT_DATE_DATE", field

	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "FORACID", field

	Set field = db.TableDef.NewField
		field.Name = "AUDIT_TIME"
		field.Description = ""
		field.Type = WI_TIME_FIELD
		field.Equation = ""
		task.ReplaceField "AUDIT_DATE_TIME", field

	Set field = db.TableDef.NewField
		field.Name = "AUTHORIZED_BY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 15
		task.ReplaceField "AUTH_ID", field

	Set field = db.TableDef.NewField
		field.Name = "REMARKS"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 100
		task.ReplaceField "RMKS", field

	Set field = db.TableDef.NewField
		field.Name = "JOINT_HOLDER_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "COMP_ACCT_SRL_NUM", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function RenameFieldsDetails
  If haveRecords("AM9_Wdrl_After_New_Joint_Dtls.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Dtls.IMD")
	Set task = db.TableManagement

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "TRAN_ID", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_TYPE"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 1
		task.ReplaceField "TRAN_TYPE", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_SUB_TYPE"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 2
		task.ReplaceField "TRAN_SUB_TYPE", field

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
		field.Name = "TRANSACTION_REMARKS"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "TRAN_RMKS", field

	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_CURRENCY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "TRANSACTION_CRNCY_CODE", field

	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "CUST_NAME", field
		
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

Function RenameFieldsSummary
If haveRecords("AM9_Wdrl_After_New_Joint_Summ.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Summ.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_NUMBER"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "FORACID", field

	Set field = db.TableDef.NewField
	field.Name = "NO_OF_WITHDRAWALS"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "ACCT_NAME", field

	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "CUST_NAME", field

	Set field = db.TableDef.NewField
	field.Name = "ACCOUNT_DESCRIPTION"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "SCHM_DESC", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExportResults
If haveRecords("AM9_Wdrl_After_New_Joint_Summ.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Summ.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "NO_OF_WITHDRAWALS"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\AM9_Wdrl_After_New_Joint_Summary.MDB", "Database", "MDB2000", 1, db.Count, eqn
RESULTSLOG(db.name)
			Set task = Nothing
	Set db = Nothing
Else 
NORESULTSLOG("AM9_SUMMARY.IMD") 
End If

If haveRecords("AM9_Joint_Parties.IMD") Then
	Set db = Client.OpenDatabase("AM9_Joint_Parties.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "JOINT_CUSTOMER_ID"
		task.AddFieldToInc "JOINT_ACCOUNT_HOLDER"
		task.AddFieldToInc "JOINT_HOLDER_NUMBER"
		task.AddFieldToInc "AUDIT_DATE"
		task.AddFieldToInc "AUDIT_TIME"
		task.AddFieldToInc "ENTERER_ID"
		task.AddFieldToInc "AUTHORIZED_BY"
		task.AddFieldToInc "REMARKS"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\AM9_Joint_Parties.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("AM9_Wdrl_After_New_Joint_Dtls.IMD") Then
	Set db = Client.OpenDatabase("AM9_Wdrl_After_New_Joint_Dtls.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "TRANSACTION_ID"
		task.AddFieldToInc "TRANSACTION_TYPE"
		task.AddFieldToInc "TRANSACTION_SUB_TYPE"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "TRANSACTION_CURRENCY"
		task.AddFieldToInc "TRANSACTION_AMOUNT"
		task.AddFieldToInc "TRANSACTION_PARTICULARS"
		task.AddFieldToInc "TRANSACTION_REMARKS"
		task.AddFieldToInc "SOL_ID"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\AM9_Wdrl_After_New_Joint_Detail.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End If	
End Function

'Delete files created to execute this script.
Function CleanUp
	DeleteFile("AAS_Records_Added.IMD")
	DeleteFile("Sorted AAS Records Added.IMD")
	DeleteFile("Joint_Parties_Added_Today.IMD")
	DeleteFile("Joint Parties Today - Main Acct Join.IMD")
	DeleteFile("Joint Parties Today - Main Acct CGM.IMD")
	DeleteFile("Joint Parties Today Complete.IMD")
	'DeleteFile("Customer Induced Txns.IMD.IMD")
	DeleteFile("Join CustTxns NewJointParties.IMD")
	DeleteFile("Wdrl_After_New_Joint_Party.IMD")
	DeleteFile("Joint Parties Today Complete - GAM.IMD")
	'DeleteFile("AM9_Wdrl_After_New_Joint_Dtls.IMD")
	'DeleteFile("AM9_Wdrl_After_New_Joint_Summ.IMD")
	'DeleteFile("AM9_Joint_Parties.IMD")
	DeleteFile("AM9_Wdrl_After_New_Joint_Summ_tmp.IMD")
	DeleteFile("AM9_Joint_Parties_tmp.IMD")
	DeleteFile("Summ AAS Records Added.IMD")
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

'---------------------------------------------------------------------------------------------------------------------------------------
' Delete files
'---------------------------------------------------------------------------------------------------------------------------------------
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

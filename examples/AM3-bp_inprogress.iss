'====================================================================================================
'	Test#: 		AM3 - Dormant account with Txns
'	Risk:		Dormant accounts may have transactions.
' 	Objective:	Identify customer induced transactions on dormant accounts.
' 	Frequency:	Daily
' 	Last Modified:	13/11/2014
'====================================================================================================
'	Script Dependencies: Import, interim
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "AM3 - Dormant account with Txns"
Const scriptname_log ="AM3 - Dormant account with Txns.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main	

	Ignorewarning(True)
	
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Client.CloseAll
	CleanUp
	
	Get_Dormant_Txn_Acct_No_Status_Change_Today	
	Join_Transaction_Status
	Comp_Dormant_Tran_FLAG
	Get_Dormant_Txn_Acct_Status_Change_Today
	Get_Final_Dormant_Txn_File
	Get_Summary_Detail
	Rename_Fields_Summary
	Rename_Fields_Details
	Join_RiskRating
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


	'Client.CloseAll
	'CleanUp
	'Client.Quit
End Sub

'=======================================================================================
'The Approach: process in parts by:
'	A. Isolating accounts that saw a change in the status today to determine if the transaction was done 
'	     while the acount was in a dormant state
'	B. Isolate all other dormant accounts that had a transaction passing on the account today
'=======================================================================================
Function Get_Dormant_Txn_Acct_No_Status_Change_Today
If haveRecords("Customer Induced Txns.IMD") Then
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "TRAN_ID"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "TRANSACTION_AMT"
			task.AddFieldToInc "TRAN_PARTICULAR"
			task.AddFieldToInc "ENTRY_USER_ID"
			task.AddFieldToInc "PSTD_USER_ID"
			task.AddFieldToInc "RCRE_USER_ID"
			task.AddFieldToInc "RCRE_TIME_DATE"
			task.AddFieldToInc "RCRE_TIME_TIME"
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddFieldToInc "PART_TRANSACTION_TYPE"
			dbName = "Customer Induced Txns Extract.IMD"
			task.AddExtraction dbName, "", ""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Customer Induced Txns Extract.IMD") And haveRecords("Account_Status_Change.IMD") Then	
	Set db = Client.OpenDatabase("Customer Induced Txns Extract.IMD")
		Set task = db.JoinDatabase
		task.FileToJoin "Account_Status_Change.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "ACID"
			task.AddMatchKey "ACID", "TABLE_KEY", "A"
			task.CreateVirtualDatabase = False
			dbName = "Cust Txn No Status Change.IMD"
			task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
		Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("TBAADM.SBCA_MAST_TABLE.IMD")  Then
	Set db = Client.OpenDatabase("TBAADM.SBCA_MAST_TABLE.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "ACCT_STATUS"
			task.AddFieldToInc "LCHG_TIME_DATE"
			task.AddFieldToInc "LCHG_TIME_TIME"
			dbName = "SBCA Dormant.IMD"
			task.AddExtraction dbName, "", "ACCT_STATUS == ""D"""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Cust Txn No Status Change.IMD") And haveRecords("SBCA Dormant.IMD") Then	
	Set db = Client.OpenDatabase("Cust Txn No Status Change.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "SBCA Dormant.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "ACCT_STATUS"
			task.AddSFieldToInc "LCHG_TIME_DATE"
			task.AddSFieldToInc "LCHG_TIME_TIME"
			task.AddMatchKey "ACID", "ACID", "A"
			task.CreateVirtualDatabase = False
			dbName = "Cust Txn Dormant Interim1.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
End Function

' Isolate transactions that occured on accounts with status changes.
Function Join_Transaction_Status
If haveRecords("Customer Induced Txns.IMD") And haveRecords("Account_Status_Change.IMD") Then	
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
		Set task = db.JoinDatabase
		task.FileToJoin "Account_Status_Change.IMD"
			task.AddPFieldToInc "ACID"
			task.AddPFieldToInc "TRAN_ID"
			task.AddPFieldToInc "TRANSACTION_DATE"
			task.AddPFieldToInc "TRANSACTION_AMT"
			task.AddPFieldToInc "TRAN_PARTICULAR"
			task.AddPFieldToInc "ENTRY_USER_ID"
			task.AddPFieldToInc "PSTD_USER_ID"
			task.AddPFieldToInc "RCRE_USER_ID"
			task.AddPFieldToInc "RCRE_TIME_DATE"
			task.AddPFieldToInc "RCRE_TIME_TIME"
			task.AddPFieldToInc "CUST_ID"
			task.AddPFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddPFieldToInc "PART_TRANSACTION_TYPE"
			task.AddMatchKey "ACID", "TABLE_KEY", "A"
			task.CreateVirtualDatabase = False
			dbName = "Cust Txn Status Change.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
	
' Create a many to many match of the records to be analysed.
If haveRecords("Cust Txn Status Change.IMD") And haveRecords("Account_Status_Change.IMD") Then
	Set db = Client.OpenDatabase("Cust Txn Status Change.IMD")
	Set task = db.VisualConnector
		id0 = task.AddDatabase ("Cust Txn Status Change.IMD")
		id1 = task.AddDatabase ("Account_Status_Change.IMD")
		task.MasterDatabase = id0
		task.AppendDatabaseNames = FALSE
		task.IncludeAllPrimaryRecords = TRUE
		task.AddRelation id0, "ACID", id1, "TABLE_KEY"
		task.AddFieldToInclude id0, "TRANSACTION_DATE"
		task.AddFieldToInclude id0, "TRAN_ID"
		task.AddFieldToInclude id0, "ACID"
		task.AddFieldToInclude id0, "TRANSACTION_AMT"
		task.AddFieldToInclude id0, "TRAN_PARTICULAR"
		task.AddFieldToInclude id0, "ENTRY_USER_ID"
		task.AddFieldToInclude id0, "PSTD_USER_ID"
		task.AddFieldToInclude id0, "RCRE_USER_ID"
		task.AddFieldToInclude id0, "RCRE_TIME_DATE"
		task.AddFieldToInclude id0, "RCRE_TIME_TIME"
		task.AddFieldToInclude id0, "CUST_ID"
		task.AddFieldToInclude id0, "TRANSACTION_CRNCY_CODE"
		task.AddFieldToInclude id0, "PART_TRANSACTION_TYPE"
		task.AddFieldToInclude id1, "REF_NUM"
		task.AddFieldToInclude id1, "TABLE_NAME"
		task.AddFieldToInclude id1, "TABLE_KEY"
		task.AddFieldToInclude id1, "ACID"
		task.AddFieldToInclude id1, "ENTERER_ID"
		task.AddFieldToInclude id1, "AUTH_ID"
		task.AddFieldToInclude id1, "MODIFIED_FIELDS_DATA"
		task.AddFieldToInclude id1, "AUDIT_DATE_DATE"
		task.AddFieldToInclude id1, "AUDIT_DATE_TIME"
		task.AddFieldToInclude id1, "COMP_STATUS_CHANGE"
		task.AddFieldToInclude id1, "COMP_CHANGE_FROM"
		task.AddFieldToInclude id1, "COMP_CHANGE_TO"
		task.CreateVirtualDatabase = False
		dbName = "Transaction Audit Change Int.IMD"
		task.OutputDatabaseName = dbName
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Transaction Audit Change Int.IMD") Then
	Set db = Client.OpenDatabase("Transaction Audit Change Int.IMD")
	Set task = db.Sort
		task.AddKey "TRAN_ID", "A"
		task.AddKey "RCRE_TIME_TIME", "A"
		task.AddKey "AUDIT_DATE_TIME", "A"
		dbName = "Transaction Audit Change.IMD"
		task.PerformTask dbName
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Append Field
Function Comp_Dormant_Tran_FLAG
If haveRecords("Transaction Audit Change.IMD") Then
	Set db = Client.OpenDatabase("Transaction Audit Change.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_DORMANT_TXN1"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(TRANSACTION_DATE = AUDIT_DATE_DATE .AND. (COMP_CHANGE_TO == ""D""  .AND. RCRE_TIME_TIME> AUDIT_DATE_TIME), ""Y"", ""N"") "
		field.Length = 1
		task.AppendField field
	
		Set field = db.TableDef.NewField
		field.Name = "COMP_DORMANT_TXN2"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(TRANSACTION_DATE = AUDIT_DATE_DATE .AND. (COMP_CHANGE_FROM == ""D""  .AND. RCRE_TIME_TIME < AUDIT_DATE_TIME), ""Y"", ""N"")"
		field.Length = 1
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Data: Direct Extraction
Function Get_Dormant_Txn_Acct_Status_Change_Today
If haveRecords("Transaction Audit Change.IMD") Then
	Set db = Client.OpenDatabase("Transaction Audit Change.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "TRAN_ID"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "TRANSACTION_AMT"
			task.AddFieldToInc "TRAN_PARTICULAR"
			task.AddFieldToInc "ENTRY_USER_ID"
			task.AddFieldToInc "PSTD_USER_ID"
			task.AddFieldToInc "RCRE_USER_ID"
			task.AddFieldToInc "RCRE_TIME_DATE"
			task.AddFieldToInc "RCRE_TIME_TIME"
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddFieldToInc "PART_TRANSACTION_TYPE"
			dbName = "Cust Txn Dormant Interim2.IMD"
			task.AddExtraction dbName, "", "COMP_DORMANT_TXN1 == ""Y""  .OR. COMP_DORMANT_TXN2 == ""Y"""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
End If
End Function


Function Get_Final_Dormant_Txn_File
  Client.CloseAll

  Dim FileOneCount As Long
  Dim FileTwoCount As Long

' Count records in both files
  If haveRecords("Cust Txn Dormant Interim2.IMD") Then
  Set db1 = Client.OpenDatabase("Cust Txn Dormant Interim2.IMD")	'---File One
         FileOneCount = db1.count
       Else
         FileOneCount = 0
  End If
  
  If haveRecords("Cust Txn Dormant Interim1.IMD") Then
  Set db2 = Client.OpenDatabase("Cust Txn Dormant Interim1.IMD")	'---File Two
     FileTwoCount = db2.count
     Else
              FileTwoCount = 0
  End If
     
  Set db1 = Nothing
  Set db2 = Nothing

             
' Go to end of function if both files are empty
If FileOneCount = 0 And FileTwoCount  = 0 Then 
   Exit Sub
 End If
 
' Test files for records before appending
      If FileOneCount > 0 And FileTwoCount  = 0 Then
        	Client.CloseAll
        	Set ProjectManagement = client.ProjectManagement
	  ProjectManagement.RenameDatabase "Cust Txn Dormant Interim2.IMD", "All Transactions"
	Set ProjectManagement = Nothing
		
       ElseIf FileOneCount = 0 And FileTwoCount  > 0 Then
         	Client.CloseAll
  	Set ProjectManagement = client.ProjectManagement
	  ProjectManagement.RenameDatabase "Cust Txn Dormant Interim1.IMD", "All Transactions"
	Set ProjectManagement = Nothing
	
       Else ' Append...No file is empty
       	Set db = Client.OpenDatabase("Cust Txn Dormant Interim1.IMD")
	Set task = db.AppendDatabase
		task.AddDatabase "Cust Txn Dormant Interim2.IMD"
		dbName = "All Transactions.IMD"
		task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
    End If

  Set db1 = Nothing
  Set db2 = Nothing
End Function

' File: Join Databases
Function Get_Summary_Detail
If haveRecords("All Transactions.IMD") And haveRecords("General Acct Master Lite.IMD")  Then
	Set db = Client.OpenDatabase("All Transactions.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "General Acct Master Lite.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "FORACID"
			task.AddSFieldToInc "SOL_ID"
			task.AddSFieldToInc "ACCT_NAME"
			task.AddSFieldToInc "CUST_NAME"
			task.AddSFieldToInc "SCHM_DESC"
			task.AddMatchKey "ACID", "ACID", "A"
			task.CreateVirtualDatabase = False
			dbName = "AM3 Details1.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
 
If haveRecords("AM3 Details1.IMD") And haveRecords("NCB_BRANCHES.IMD")  Then
	Set db = Client.OpenDatabase("AM3 Details1.IMD")
		Set task = db.JoinDatabase
		   	task.FileToJoin "NCB_BRANCHES.IMD"
		   	task.IncludeAllPFields
		   	task.AddSFieldToInc "BR_NAME"
		   	task.AddMatchKey "SOL_ID", "MICR_BRANCH_CODE", "A"
			task.CreateVirtualDatabase = False
			dbName = "AM3 Details Branch.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
 
'Get current account status
If haveRecords("AM3 Details Branch.IMD") And HaveRecords("TBAADM.SBCA_MAST_TABLE.IMD") Then
 	Set db = Client.OpenDatabase("AM3 Details Branch.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "TBAADM.SBCA_MAST_TABLE.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "ACCT_STATUS"
			task.AddMatchKey "ACID", "ACID", "A"
			task.CreateVirtualDatabase = False
			dbName = "AM3 Details.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("AM3 Details.IMD") Then
	Set db = Client.OpenDatabase("AM3 Details.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "ACID"
		task.AddFieldToInc "CUST_ID"
		task.AddFieldToInc "FORACID"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "BR_NAME"
		task.AddFieldToInc "ACCT_NAME"
		task.AddFieldToInc "ACCT_STATUS1"
		task.AddFieldToInc "CUST_NAME"
		task.AddFieldToInc "SCHM_DESC"
		task.AddFieldToInc "TRANSACTION_DATE"
		dbName = "AM3 Summary_tmp.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Modify Field

Function Rename_Fields_Summary
If haveRecords("AM3 Summary_tmp.IMD") Then
	Set db = Client.OpenDatabase("AM3 Summary_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_NUMBER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "FORACID", field
		
	Set field = db.TableDef.NewField
		field.Name = "BRANCH_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "BR_NAME", field

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
		field.Name = "ACCOUNT_DESCRIPTION"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 25
		task.ReplaceField "SCHM_DESC", field
		
	Set field = db.TableDef.NewField
		field.Name = "CURRENT_STATUS"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 1
		task.ReplaceField "ACCT_STATUS1", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function Rename_Fields_Details
If haveRecords("AM3 Details.IMD") Then
	Set db = Client.OpenDatabase("AM3 Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_NAME"
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
		
	Set field = db.TableDef.NewField
		field.Name = "BRANCH_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "BR_NAME", field
		task.PerformTask

	Set field = db.TableDef.NewField
		field.Name = "TRANACTION_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "TRAN_ID", field

	Set field = db.TableDef.NewField
		field.Name = "CREATED_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "RCRE_TIME_DATE", field

	Set field = db.TableDef.NewField
		field.Name = "CREATED_TIME"
		field.Description = ""
		field.Type = WI_TIME_FIELD
		field.Equation = ""
		task.ReplaceField "RCRE_TIME_TIME", field
		task.PerformTask
		


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
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
	
'	Set db = Client.OpenDatabase("AM3 Details.IMD")	
'	Set task = db.TableManagement
'	Set field = db.TableDef.NewField
'	field.Name = "CREATED_TIME"
'	field.Description = ""
'	field.Type = WI_VIRT_TIME
'	field.Equation = "@Ctot(RCRE_TIME_TIME,""HH:MM:SS"")"
'	task.AppendField field
'	task.PerformTask
'	Set field = Nothing
'	Set task = Nothing
'	Set db = Nothing
End If
End Function

Function Join_RiskRating
'If haveRecords("ACID_FORACID.IMD") And haveRecords("AM3 Summary_tmp.IMD") Then
'	Set db = Client.OpenDatabase("AM3 Summary_tmp.IMD")	
'	Set task = db.JoinDatabase
'		task.FileToJoin "ACID_FORACID.IMD" 
'		task.IncludeAllPFields
'		task.AddSFieldToInc "ACID"
'		task.AddMatchKey "ACCOUNT_NUMBER", "FORACID", "A"
'		task.CreateVirtualDatabase = False
'		dbName = "AM3 Summary_tmp1.IMD"
'		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
'	Set task = Nothing
'	Set db = Nothing
'End If


If haveRecords("Account_Turnover_wRisk.IMD") And haveRecords("AM3 Summary_tmp.IMD") Then
	Set db = Client.OpenDatabase("AM3 Summary_tmp.IMD")	
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
		dbName = "AM3 Summary.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

' File - Export Database
Function ExportDatabase
If haveRecords("AM3 Summary.IMD") Then
	Set db = Client.OpenDatabase("AM3 Summary.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "CURRENT_STATUS"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "BRANCH_NAME"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\AM3_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
	RESULTSLOG(db.name)
	Set task = Nothing
	Set db = Nothing
Else 
NORESULTSLOG("AM3_SUMMARY.IMD") 
	
End If
	
If haveRecords("AM3 Details.IMD") Then
	Set db = Client.OpenDatabase("AM3 Details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "BRANCH_NAME"
		task.AddFieldToInc "TRANACTION_ID"
		task.AddFieldToInc "TRANSACTION_CURRENCY"
		task.AddFieldToInc "TRANSACTION_DATE"
		task.AddFieldToInc "TRANSACTION_AMT"
		task.AddFieldToInc "TRAN_PARTICULAR"
		task.AddFieldToInc "CREATED_DATE"
		task.AddFieldToInc "CREATED_TIME"
		task.AddFieldToInc "PART_TRANSACTION_TYPE"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\AM3_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
 
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function CleanUp
	DeleteFile("Customer Induced Txns Extract.IMD")
	DeleteFile("Cust Txn No Status Change.IMD")
	DeleteFile("SBCA Dormant.IMD")
	DeleteFile("Cust Txn Dormant Interim1.IMD")
	DeleteFile("Cust Txn Status Change.IMD")
	DeleteFile("Transaction Audit Change Int.IMD")
	DeleteFile("Transaction Audit Change.IMD")
	DeleteFile("Cust Txn Dormant Interim2.IMD")
	DeleteFile("All Transactions.IMD")
	DeleteFile("AM3 Details1.IMD")
	DeleteFile("AM3 Details Branch.IMD")
	DeleteFile("AM3 Details.IMD")
	'DeleteFile("AM3 Summary.IMD")
	DeleteFile("AM3 Summary_tmp.IMD")
	'DeleteFile("AM3 Summary_tmp1.IMD")
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

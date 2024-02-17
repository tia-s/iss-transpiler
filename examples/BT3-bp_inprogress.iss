'====================================================================================================
'	Test#: 		BT3 - Multiple Transactions Over Threshold
'	Risk:		
' 	Objective:	
' 	Frequency:	Daily
' 	Last Modified:	14/07/2016 03:42:02 PM
'	Comment:	Added ministry exceptions and updated cleanup
'
'====================================================================================================
'	Script Dependencies:
'====================================================================================================
'----- Constants -----

Const scriptname_log ="BT3 - Multiple Transactions Over Threshold.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main
	Ignorewarning(True)

	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Client.CloseAll
	CleanUp

	
	Call Add_Account_Info_to_Transaction

	Call Get_Multi_Transactions_by_Same_customer_same_Day
	
	Call Get_Multi_Transactions_Over_Threshold_Summ

	Call Get_Multi_Transactions_Over_Threshold_Details	
	
	Call Join_RiskRating
	
	Call ExportDatabase
	

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


Function Add_Account_Info_to_Transaction
If haveRecords("General Acct Master Lite.IMD")  Then
	Set db = Client.OpenDatabase("General Acct Master Lite.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "FORACID"
		task.AddFieldToInc "ACID"
		task.AddFieldToInc "ACCT_NAME"
		task.AddFieldToInc "CUST_NAME"
		task.AddFieldToInc "CUST_PERM_ADDR1"
		task.AddFieldToInc "CUST_PERM_ADDR2"
		task.AddFieldToInc "SCHM_DESC"
		task.AddFieldToInc "CUSTOMER_TYPE"
		dbName = "GAM Extract.IMD"
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("General Acct Master Lite.IMD", "Add_Account_Info_to_Transaction", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("GAM Extract.IMD") And haveRecords("Customer Induced Txns.IMD") Then
	Set db = Client.OpenDatabase("Customer Induced Txns.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "GAM Extract.IMD"
		task.IncludeAllPFields
		task.IncludeAllSFields
		task.AddMatchKey "ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		task.Criteria = "TRAN_TYPE == ""C"""
		dbName = "Customer_Induced_Trans_Cust_Acct_Info_Int.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("GAM Extract.IMD or Customer Induced Txns.IMD", "Add_Account_Info_to_Transaction", "JoinDatabase", "Error", "Databases empty or does not exist.")
	
End If

If haveRecords("Customer_Induced_Trans_Cust_Acct_Info_Int.IMD") Then
	Set db = Client.OpenDatabase("Customer_Induced_Trans_Cust_Acct_Info_Int.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUST_NUM"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@Val(CUST_ID)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If
If haveRecords("Customer_Induced_Trans_Cust_Acct_Info_Int.IMD") And haverecords("Ministries-Database.IMD") Then
	Set db = Client.OpenDatabase("Customer_Induced_Trans_Cust_Acct_Info_Int.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "Ministries-Database.IMD"
	task.IncludeAllPFields
	task.AddMatchKey "CUST_NUM", "CUST_ID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Customer_Induced_Trans_Cust_Acct_Info1.IMD"
	task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
End If		
End Function

Function Get_Multi_Transactions_by_Same_customer_same_Day
' Get Transaction branch
If haveRecords("TBAADM.DAILY_TRAN_HEADER_TABLE.IMD") And haveRecords("NCB_BRANCHES.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.DAILY_TRAN_HEADER_TABLE.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "INIT_SOL_ID"
		task.AddFieldToInc "TRAN_ID"
		dbName = "Tran Header Branch.IMD"
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.DAILY_TRAN_HEADER_TABLE.IMD or NCB_BRANCHES.IMD", "Get_Multi_Transactions_by_Same_customer_same_Day", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If
If haverecords ("Tran Header Branch.IMD")	Then
	Set db = Client.OpenDatabase("Tran Header Branch.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "NCB_BRANCHES.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "BR_NAME"
		task.AddMatchKey "INIT_SOL_ID", "MICR_BRANCH_CODE", "A"
		task.CreateVirtualDatabase = False
		dbName = "Transaction Branch.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
If haverecords ("Transaction Branch.IMD")	Then
	Set db = Client.OpenDatabase("Transaction Branch.IMD")	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_BRANCH"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 30
		task.ReplaceField "BR_NAME", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_SOL"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 8
		task.ReplaceField "INIT_SOL_ID", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If

' Exclude customers with single transactions over the threshold
If haveRecords("Customer_Induced_Trans_Cust_Acct_Info1.IMD") And haverecords("Transaction Branch.IMD") Then
	Set db = Client.OpenDatabase("Customer_Induced_Trans_Cust_Acct_Info1.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Transaction Branch.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "TRANSACTION_SOL"
		task.AddSFieldToInc "TRANSACTION_BRANCH"
		task.AddMatchKey "TRAN_ID", "TRAN_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Customer_Induced_Trans_Cust_Acct_Info.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	
	Set db = Client.OpenDatabase("Customer_Induced_Trans_Cust_Acct_Info.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "CUST_ID"
		dbName = "Single Txn Above Threshold.IMD"
		'task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US >=" & e_TTR_Value_US
		task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US >= 15000" '- EDITED BY A.M. TO BE CHANGED BACK
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Customer_Induced_Trans_Cust_Acct_Info1.IMD or Transaction Branch.IMD", "Get_Multi_Transactions_by_Same_customer_same_Day", "JoinDatabase", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("Customer_Induced_Trans_Cust_Acct_Info.IMD") And haveRecords("Single Txn Above Threshold.IMD") Then
	Set db = Client.OpenDatabase("Customer_Induced_Trans_Cust_Acct_Info.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Single Txn Above Threshold.IMD"
		task.AddPFieldToInc "CUST_ID"
		task.AddPFieldToInc "TRAN_PARTICULAR"
		task.AddPFieldToInc "TRANSACTION_DATE"
		task.AddPFieldToInc "TRANSACTION_SOL"
		task.AddPFieldToInc "TRANSACTION_BRANCH"
		task.AddPFieldToInc "TRAN_ID"
		task.AddPFieldToInc "TRAN_TYPE"
		task.AddPFieldToInc "TRAN_SUB_TYPE"
		task.AddPFieldToInc "PART_TRAN_TYPE"
		task.AddPFieldToInc "PART_TRANSACTION_TYPE"
		task.AddPFieldToInc "ACID"
		task.AddPFieldToInc "VALUE_DATE_DATE"
		task.AddPFieldToInc "TRANSACTION_AMT"
		task.AddPFieldToInc "TRANSACTION_CRNCY_CODE"
		task.AddPFieldToInc "FX_RATE_TO_JMD"
		task.AddPFieldToInc "USD"
		task.AddPFieldToInc "TRANSACTION_AMOUNT_JM"
		task.AddPFieldToInc "FX_RATE_JMD_TO_USD"
		task.AddPFieldToInc "TRANSACTION_AMOUNT_US"
		task.AddPFieldToInc "BR_NAME"
		task.AddPFieldToInc "SOL_ID"
		task.AddPFieldToInc "FORACID"
		task.AddPFieldToInc "ACCT_NAME"
		task.AddPFieldToInc "SCHM_DESC"
		task.AddPFieldToInc "CUST_NAME"
		task.AddPFieldToInc "CUST_PERM_ADDR1"
		task.AddPFieldToInc "CUST_PERM_ADDR2"
		task.AddPFieldToInc "CUSTOMER_TYPE"
		task.AddSFieldToInc "CUST_ID"
		task.AddMatchKey "CUST_ID", "CUST_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Customers Without Single Txn Over.IMD"
		task.PerformTask dbName, "", WI_JOIN_NOC_SEC_MATCH 'WI_JOIN_ALL_IN_PRIM - EDITED BY A.M. TO BE CHANGED BACK TO WI_JOIN_NOC_SEC_MATCH
	Set task = Nothing
	Set db = Nothing
	
Else
'Rename original file (Customer_Induced_Trans_Cust_Acct_Info) if there are no single transactions over threshold.
	If haveRecords("Customer_Induced_Trans_Cust_Acct_Info.IMD") Then
		Set ProjectManagement = client.ProjectManagement
			ProjectManagement.RenameDatabase "Customer_Induced_Trans_Cust_Acct_Info.IMD", "Customers Without Single Txn Over"
		Set ProjectManagement = Nothing
	End If
End If

' Summarize transactions to get total and number of transactions.
If haveRecords("Customers Without Single Txn Over.IMD") Then
	Set db = Client.OpenDatabase("Customers Without Single Txn Over.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "CUST_ID"
	task.AddFieldToSummarize "TRANSACTION_DATE"
	task.AddFieldToSummarize "PART_TRANSACTION_TYPE"
	task.AddFieldToInc "CUST_NAME"
	task.AddFieldToInc "CUST_PERM_ADDR1"
	task.AddFieldToInc "CUST_PERM_ADDR2"
	task.AddFieldToInc "CUSTOMER_TYPE"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_JM"
	task.AddFieldToTotal "TRANSACTION_AMOUNT_US"
	dbName = "Multi_Transactions_by_Same_customer_same_Day.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Customers Without Single Txn Over.IMD", "Get_Multi_Transactions_by_Same_customer_same_Day", "Summarization", "Error", "Databases empty or does not exist.")		
	
End If
End Function

Function Get_Multi_Transactions_Over_Threshold_Summ
If haveRecords("Multi_Transactions_by_Same_customer_same_Day.IMD") Then
	Set db = Client.OpenDatabase("Multi_Transactions_by_Same_customer_same_Day.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "BT3_Multi_Transactions_Over_Threshold_Summ1.IMD"
	'task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US_SUM >= " & e_TTR_Value_US & " .AND. NO_OF_RECS >=" & e_Minimum_Transactions_Count
	task.AddExtraction dbName, "", "TRANSACTION_AMOUNT_US_SUM >= 15000 .AND. NO_OF_RECS >= 2" '- EDITED BY A.M. TO BE CHANGED BACK
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Multi_Transactions_by_Same_customer_same_Day.IMD", "Get_Multi_Transactions_Over_Threshold_Summ", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ1.IMD") And haveRecords("Customer Master Lite.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Summ1.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "Customer Master Lite.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "PRIMARY_SOL_ID"
			task.AddSFieldToInc "PAN_GIR_NUM"
			task.AddMatchKey "CUST_ID", "CUST_ID", "A"
			task.CreateVirtualDatabase = False
			dbName = "BT3_Multi_Transactions_Over_Threshold_Summ2.IMD"
			task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("BT3_Multi_Transactions_Over_Threshold_Summ1.IMD or Customer Master Lite.IMD", "Get_Multi_Transactions_Over_Threshold_Summ", "JoinDatabase", "Error", "Databases empty or does not exist.")	
	
End If

If haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ2.IMD") And haveRecords("NCB_BRANCHES.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Summ2.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "NCB_BRANCHES.IMD"
	task.AddPFieldToInc "CUST_ID"
	task.AddPFieldToInc "CUST_NAME"
	task.AddPFieldToInc "TRANSACTION_DATE"
	task.AddPFieldToInc "NO_OF_RECS"
	task.AddPFieldToInc "PAN_GIR_NUM"
	task.AddPFieldToInc "CUST_PERM_ADDR1"
	task.AddPFieldToInc "CUST_PERM_ADDR2"
	task.AddPFieldToInc "CUSTOMER_TYPE"
	task.AddPFieldToInc "TRANSACTION_AMOUNT_JM_SUM"
	task.AddPFieldToInc"TRANSACTION_AMOUNT_US_SUM"
	task.AddPFieldToInc "PRIMARY_SOL_ID"
	task.AddPFieldToInc "PART_TRANSACTION_TYPE"
	task.AddSFieldToInc "BR_NAME"
	task.AddMatchKey "PRIMARY_SOL_ID", "MICR_BRANCH_CODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("BT3_Multi_Transactions_Over_Threshold_Summ2.IMD or NCB_BRANCHES.IMD", "Get_Multi_Transactions_Over_Threshold_Summ", "JoinDatabase", "Error", "Databases empty or does not exist.")	
	
End If

	
If haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 12
	task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField	
	field.Name = "BRANCH_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 30
	task.ReplaceField "BR_NAME", field
	
	Set field = db.TableDef.NewField
	field.Name = "PERMANENT_ADDRESS1"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 45
	task.ReplaceField "CUST_PERM_ADDR1", field

	Set field = db.TableDef.NewField
	field.Name = "PERMANENT_ADDRESS2"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 45
	task.ReplaceField "CUST_PERM_ADDR2", field

	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_TRN"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "PAN_GIR_NUM", field
		
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_TRANSACTIONS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	
	field.Name = "CUSTOMER_BRANCH_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 50
	task.ReplaceField "BR_NAME", field

	Set field = db.TableDef.NewField
	field.Name = "TOTAL_TRANSACTION_AMOUNT_JM"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "TRANSACTION_AMOUNT_JM_SUM", field

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
	field.Length = 120
	task.ReplaceField "CUST_NAME", field
	
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Get_Multi_Transactions_Over_Threshold_Details
If haveRecords("Customers Without Single Txn Over.IMD") And haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD")   Then
	Set db = Client.OpenDatabase("Customers Without Single Txn Over.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD"
		task.AddPFieldToInc "TRAN_ID"
		task.AddPFieldToInc "TRANSACTION_DATE"
		task.AddPFieldToInc "TRANSACTION_AMT"
		task.AddPFieldToInc "TRANSACTION_CRNCY_CODE"
		task.AddPFieldToInc "TRAN_PARTICULAR"
		task.AddPFieldToInc "FX_RATE_TO_JMD"
		task.AddPFieldToInc "TRANSACTION_AMOUNT_JM"
		task.AddPFieldToInc "FX_RATE_JMD_TO_USD"
		task.AddPFieldToInc "TRANSACTION_AMOUNT_US"
		task.AddPFieldToInc "PART_TRANSACTION_TYPE"
		task.AddPFieldToInc "TRANSACTION_SOL"
		task.AddPFieldToInc "TRANSACTION_BRANCH"
		task.AddPFieldToInc "FORACID"
		task.AddPFieldToInc "ACCT_NAME"
		task.AddPFieldToInc "SOL_ID"
		task.AddPFieldToInc "BR_NAME"
		task.AddPFieldToInc "SCHM_DESC"
		task.AddSFieldToInc "CUSTOMER_ID"
		task.AddSFieldToInc "CUSTOMER_NAME"
		task.AddSFieldToInc "CUSTOMER_TRN"
		task.AddSFieldToInc "PERMANENT_ADDRESS1"
		task.AddSFieldToInc "PERMANENT_ADDRESS2"
		task.AddSFieldToInc "CUSTOMER_TYPE"
		task.AddMatchKey "CUST_ID", "CUSTOMER_ID", "A"
		task.AddMatchKey "PART_TRANSACTION_TYPE", "PART_TRANSACTION_TYPE", "A"
		task.CreateVirtualDatabase = False
		dbName = "BT3_Multi_Transactions_Over_Threshold_Details.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Customers Without Single Txn Over.IMD or BT3_Multi_Transactions_Over_Threshold_Summ.IMD", "Get_Multi_Transactions_Over_Threshold_Details", "JoinDatabase", "Error", "Databases empty or does not exist.")	
	
End If
	
	
If haveRecords("BT3_Multi_Transactions_Over_Threshold_Details.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField

	
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
	field.Name = "ACCOUNT_DESC"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "SCHM_DESC", field	

	Set field = db.TableDef.NewField	
	field.Name = "BRANCH_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 30
	task.ReplaceField "BR_NAME", field
	
	Set field = db.TableDef.NewField	
	field.Name = "PAYEE_DETAILS "
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "TRAN_PARTICULAR", field
		
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "TRAN_ID", field	
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Join_RiskRating
If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD")	
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
		dbName = "BT3_Multi_Transactions_Over_Threshold_Summ.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExportDatabase
If haveRecords("BT3_Multi_Transactions_Over_Threshold_Details.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Details.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\BT3_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End If

If haveRecords("BT3_Multi_Transactions_Over_Threshold_Summ.IMD") Then
	Set db = Client.OpenDatabase("BT3_Multi_Transactions_Over_Threshold_Summ.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask Client.WorkingDirectory & "Reports\BT3_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
RESULTSLOG(db.name)
	Set db = Nothing
	Set task = Nothing
	
Else 
NORESULTSLOG("BT3_SUMMARY.IMD") 
End If


End Function



Function CleanUp
	DeleteFile("Customer_Induced_Trans_Cust_Acct_Info_Int.IMD") 

	DeleteFile("gam extract.IMD")
	DeleteFile("Tran Header Branch.IMD")
	DeleteFile("Transaction Branch.IMD")
	DeleteFile("Customer_Induced_Trans_Cust_Acct_Info.IMD")
	DeleteFile("Customer_Induced_Trans_Cust_Acct_Info1.IMD")
	DeleteFile("Single Txn Above Threshold.IMD")
	DeleteFile("Customers Without Single Txn Over.IMD")
	DeleteFile("Multi_Transactions_by_Same_customer_same_Day.IMD") 
	'DeleteFile("BT3_Multi_Transactions_Over_Threshold_Summ.IMD")
	DeleteFile("BT3_Multi_Transactions_Over_Threshold_Summ_tmp.IMD")
	'DeleteFile("BT3_Multi_Transactions_Over_Threshold_Details.IMD")
	DeleteFile("BT3_Multi_Transactions_Over_Threshold_Summ1.IMD")
	DeleteFile("BT3_Multi_Transactions_Over_Threshold_Summ2.IMD")
	DeleteFile("BT3_Multi_Transactions_Over_Threshold_Summ3.IMD")
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




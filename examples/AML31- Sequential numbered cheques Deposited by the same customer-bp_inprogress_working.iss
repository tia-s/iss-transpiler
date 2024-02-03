'====================================================================================================
'	Test#: 		AML31 - Sequential numbered cheques Deposited by the same customer
'	Risk:		Customer may deposit sequential numbered cheques.
' 	Objective:	Identify deposit of sequential numbered cheques by the same customer within a week.
' 	Frequency:	Weekly
' 	Last Modified:	14-Nov-2014
'====================================================================================================
'	Script Dependencies: Import, Interim
'====================================================================================================

'----- Constants -----
Const scriptname_log ="AML31 - Sequential numbered cheques Deposited by the same customer.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main
 	Ignorewarning(True)
	Client.CloseAll
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

 
 
	CleanUp

	Extract_Active_OCP_OCI
	Get_Clearing_Cheques
	Get_Regular_Cheques
	Client.CloseAll
	Get_Foreign_Cheques
	Get_Final_Cheque_File

	Client.CloseAll
	Get_Multiple_Cheques_deposited	
	Get_Sequential_Cheque_Deposited
	Get_Customer_Info
	
	Client.CloseAll
	Get_Multiple_Cheques_deposited_Details
	Get_Multiple_Cheques_deposited_Summ
	Join_RiskRating
	Client.CloseAll
	
	Call Export_to_MDB("AML31_Multiple_Cheques_deposited_to_same_Account_Summ", "")
	Call Export_to_MDB("AML31_Multiple_Cheques_deposited_to_same_Account_Details", "")
	
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
'	Client.Quit
End Sub



'-------------------------------------------------------------------------------------------------------------------------------
'================================ ANALYSIS ================================
'--------------------------------------------------------------------------------------------------------------------------------
'********* Get foreign check details to be included in analysis

' Get foreign check bill ids
Function Get_Foreign_Cheques
If haveRecords("TBAADM.FX_BILL_MASTER_TABLE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.FX_BILL_MASTER_TABLE.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "Foreign Cheques.IMD"
		task.AddExtraction dbName, "", "REG_TYPE ==""ICQFI"" .AND. REG_SUB_TYPE == ""CHQ"" .AND. DEL_FLG ==""N"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.FX_BILL_MASTER_TABLE.IMD", "Get_Foreign_Cheques", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

' Get foreign check transaction details
If haveRecords("TBAADM.FEX_ACCT_ENTRY_TABLE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.FEX_ACCT_ENTRY_TABLE.IMD")
	Set task = db.Extraction
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "BILL_ID"
		task.AddFieldToInc "TRAN_ID"
		task.AddFieldToInc "TRAN_DATE_DATE"
		task.AddFieldToInc "PART_TRAN_SRL_NUM"
		task.AddFieldToInc "BILL_FUNC"
		task.AddFieldToInc "TRAN_TYPE"
		task.AddFieldToInc "PART_TRAN_TYPE"
		task.AddFieldToInc "ACID"
		task.AddFieldToInc "TRAN_PARTICULAR"
		task.AddFieldToInc "TRAN_CRNCY_CODE"
		task.AddFieldToInc "TRAN_PARTICULAR_CODE"
		dbName = "FEX ACCT ENTRY EXTRACT.IMD"
		task.AddExtraction dbName, "", "PART_TRAN_TYPE == ""D"" .AND. DEL_FLG ==""N"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.FEX_ACCT_ENTRY_TABLE.IMD", "Get_Foreign_Cheques", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

' Get foriegn check numbers - all numbers for bill items (foreign checks, wire, drafts) are stored in FCIT.
If haveRecords("TBAADM.FEX_CLEAN_INST_TABLE.IMD") And  haveRecords("FEX ACCT ENTRY EXTRACT.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.FEX_CLEAN_INST_TABLE.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "FEX ACCT ENTRY EXTRACT.IMD"
		task.AddPFieldToInc "SRL_NUM"
		task.AddPFieldToInc "DEL_FLG"
		task.AddPFieldToInc "INSTRMNT_NUM"
		task.AddPFieldToInc "NUM_OF_INSTRMNT"
		task.AddPFieldToInc "INSTRMNT_TYPE"
		task.AddPFieldToInc "INSTRMNT_AMT"
		task.AddPFieldToInc "INSTRMNT_DATE_DATE"
		task.IncludeAllSFields
		task.AddMatchKey "BILL_ID", "BILL_ID", "A"
		task.AddMatchKey "SOL_ID", "SOL_ID", "A"
		task.Criteria = "DEL_FLG ==""N"""
		task.CreateVirtualDatabase = False
		dbName = "Instrument Acct Entry.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.FEX_CLEAN_INST_TABLE.IMD or FEX ACCT ENTRY EXTRACT.IMD", "Get_Foreign_Cheques", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If

' Get transaction branch details
If haveRecords("Instrument Acct Entry.IMD")  And haveRecords("NCB_BRANCHES.IMD") Then
	Set db = Client.OpenDatabase("Instrument Acct Entry.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "NCB_BRANCHES.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "BR_NAME"
		task.AddMatchKey "SOL_ID", "MICR_BRANCH_CODE", "A"
		task.CreateVirtualDatabase = False
		dbName = "Instrument Acct Branch.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Instrument Acct Entry.IMD or NCB_BRANCHES.IMD", "Get_Foreign_Cheques", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If
	
If haveRecords("Instrument Acct Branch.IMD")  And haveRecords("Foreign Cheques.IMD") Then
	Set db = Client.OpenDatabase("Instrument Acct Branch.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Foreign Cheques.IMD"
		task.AddPFieldToInc "INSTRMNT_AMT"
		task.AddPFieldToInc "INSTRMNT_NUM"
		task.AddPFieldToInc "TRAN_CRNCY_CODE"
		task.AddPFieldToInc "BR_NAME"		
		task.AddPFieldToInc "TRAN_PARTICULAR"
		task.AddPFieldToInc "TRAN_TYPE"
		task.AddPFieldToInc "TRAN_DATE_DATE"
		task.AddPFieldToInc "TRAN_ID"		
		task.AddSFieldToInc "PARTY_CODE"
		task.AddSFieldToInc "OPER_ACID"
		task.AddMatchKey "BILL_ID", "BILL_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Foreign Cheque Details.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Instrument Acct Branch.IMD or Foreign Cheques.IMD", "Get_Foreign_Cheques", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If

' Prepare fields to be included in the append to the other cheque files.
If haveRecords("Foreign Cheque Details.IMD")  Then	
	Set db = Client.OpenDatabase("Foreign Cheque Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_CRNCY_CODE"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "TRAN_CRNCY_CODE", field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "TRAN_DATE_DATE", field
	
	Set field = db.TableDef.NewField
		field.Name = "CLG_ZONE_DATE_DATE"
		field.Description = ""
		field.Type = WI_VIRT_DATE
		field.Equation = "@CTOD(""00000000"", ""YYYYMMDD"")"
		task.AppendField field

	Set field = db.TableDef.NewField
		field.Name = "TRANSACTION_AMT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "INSTRMNT_AMT", field

	Set field = db.TableDef.NewField
		field.Name = "CLG_ZONE_CODE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """"""
		field.Length = 2
		task.AppendField field

	Set field = db.TableDef.NewField
		field.Name = "CUST_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "PARTY_CODE", field

	Set field = db.TableDef.NewField
		field.Name = "ACID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 11
		task.ReplaceField "OPER_ACID", field
		task.PerformTask
	Set task = Nothing
	Set field = Nothing
	Set db = Nothing
End If
End Function

Function Extract_Active_OCP_OCI
If haveRecords("TBAADM.OUT_CLG_PART_TRAN_TABLE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.OUT_CLG_PART_TRAN_TABLE.IMD")
		Set task = db.Extraction
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "CLG_ZONE_DATE_DATE"
		task.AddFieldToInc "CLG_ZONE_CODE"
		task.AddFieldToInc "SET_NUM"
		task.AddFieldToInc "PART_TRAN_SRL_NUM"
		task.AddFieldToInc "STATUS_FLG"
		task.AddFieldToInc "TRAN_DATE_DATE"
		task.AddFieldToInc "TRAN_ID"
		dbName = "Active Regularised OCP.IMD"
		task.AddExtraction dbName, "", "DEL_FLG == ""N"" .AND. STATUS_FLG == ""G"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.OUT_CLG_PART_TRAN_TABLE.IMD", "Extract_Active_OCP_OCI", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If
	
If haveRecords("TBAADM.OUT_CLG_INSTRMNT_TABLE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.OUT_CLG_INSTRMNT_TABLE.IMD")
		Set task = db.Extraction
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "CLG_ZONE_DATE_DATE"
		task.AddFieldToInc "INSTRMNT_ID"
		task.AddFieldToInc "CLG_ZONE_CODE"
		task.AddFieldToInc "SET_NUM"
		task.AddFieldToInc "INSTRMNT_SRL_NUM"
		task.AddFieldToInc "INSTRMNT_DATE_DATE"
		task.AddFieldToInc "INSTRMNT_AMT"
		task.AddFieldToInc "BANK_CODE"
		task.AddFieldToInc "BR_CODE"
		task.AddFieldToInc "STATUS_FLG"
		task.AddFieldToInc "ENTRY_USER_ID"
		task.AddFieldToInc "ENTRY_DATE_DATE"
		task.AddFieldToInc "PAYING_ACCT_ID"
		task.AddFieldToInc "CRNCY_CODE"
		dbName = "Active Regularised OCI.IMD"
		task.AddExtraction dbName, "", "DEL_FLG == ""N"" .AND. STATUS_FLG == ""G"""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.OUT_CLG_INSTRMNT_TABLE.IMD", "Extract_Active_OCP_OCI", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If
End Function

' File: Join Databases
Function Get_Clearing_Cheques
If haveRecords("Cheque Txns.IMD") And haveRecords("Active Regularised OCP.IMD")  Then
	Set db = Client.OpenDatabase("Cheque Txns.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "Active Regularised OCP.IMD"
			task.AddPFieldToInc "TRAN_PARTICULAR"
			task.AddPFieldToInc "TRAN_TYPE"
			task.AddPFieldToInc "TRANSACTION_DATE"
			task.AddPFieldToInc "TRAN_ID"
			task.AddPFieldToInc "ACID"
			task.AddPFieldToInc "CUST_ID"
			task.AddPFieldToInc "BR_NAME"
			task.AddSFieldToInc "CLG_ZONE_DATE_DATE"
			task.AddSFieldToInc "CLG_ZONE_CODE"
			task.AddSFieldToInc "SET_NUM"
			task.AddSFieldToInc "SOL_ID"
			task.AddMatchKey "SOL_ID", "SOL_ID", "A"
			task.AddMatchKey "TRANSACTION_DATE", "TRAN_DATE_DATE", "A"
			task.AddMatchKey "PART_TRAN_SRL_NUM", "PART_TRAN_SRL_NUM", "A"
			task.AddMatchKey "TRAN_ID", "TRAN_ID", "A"
			task.CreateVirtualDatabase = False
			dbName = "HTD OCP.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Cheque Txns.IMD or Active Regularised OCP.IMD", "Get_Clearing_Cheques", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If
	
If haveRecords("Active Regularised OCI.IMD") And haveRecords("HTD OCP.IMD") Then
	Set db = Client.OpenDatabase("Active Regularised OCI.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "HTD OCP.IMD"
			task.AddSFieldToInc "BR_NAME"
			task.AddSFieldToInc "TRAN_PARTICULAR"
			task.AddSFieldToInc "TRAN_TYPE"
			task.AddSFieldToInc "TRANSACTION_DATE"
			task.AddSFieldToInc "TRAN_ID"
			task.AddSFieldToInc "ACID"
			task.AddSFieldToInc "CUST_ID"
			task.AddPFieldToInc "INSTRMNT_AMT"
			task.AddPFieldToInc "INSTRMNT_ID"
			task.AddPFieldToInc "CLG_ZONE_CODE"
			task.AddPFieldToInc "CLG_ZONE_DATE_DATE"
			task.AddPFieldToInc "CRNCY_CODE"
			task.AddMatchKey "SOL_ID", "SOL_ID", "A"
			task.AddMatchKey "CLG_ZONE_DATE_DATE", "CLG_ZONE_DATE_DATE", "A"
			task.AddMatchKey "CLG_ZONE_CODE", "CLG_ZONE_CODE", "A"
			task.AddMatchKey "SET_NUM", "SET_NUM", "A"
			task.CreateVirtualDatabase = False
			dbName = "Clearing Cheques.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Active Regularised OCI.IMD or HTD OCP.IMD", "Get_Clearing_Cheques", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If

If haveRecords("Clearing Cheques.IMD") Then
	Set db = Client.OpenDatabase("Clearing Cheques.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "INSTRMNT_NUM"
				field.Description = ""
				field.Type = WI_CHAR_FIELD
				field.Equation = ""
				field.Length = 16
				task.ReplaceField "INSTRMNT_ID", field
								
			Set field = db.TableDef.NewField
				field.Name = "TRANSACTION_CRNCY_CODE"
				field.Description = ""
				field.Type = WI_CHAR_FIELD
				field.Equation = ""
				field.Length = 3
				task.ReplaceField "CRNCY_CODE", field
				
			Set field = db.TableDef.NewField
				field.Name = "TRANSACTION_AMT"
				field.Description = ""
				field.Type = WI_NUM_FIELD
				field.Equation = ""
				field.Decimals = 4
				task.ReplaceField "INSTRMNT_AMT", field
				task.PerformTask
			Set field = Nothing
		Set task = Nothing
	Set db = Nothing
End If
End Function
	
Function Get_Regular_Cheques
If haveRecords("Cheque Txns.IMD") Then
	Set db = Client.OpenDatabase("Cheque Txns.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "BR_NAME"
			task.AddFieldToInc "TRAN_PARTICULAR"
			task.AddFieldToInc "TRAN_TYPE"
			task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
			task.AddFieldToInc "TRANSACTION_DATE"
			task.AddFieldToInc "TRAN_ID"
			task.AddFieldToInc "ACID"
			task.AddFieldToInc "TRANSACTION_AMT"
			task.AddFieldToInc "INSTRMNT_NUM"
			task.AddFieldToInc "CUST_ID"
			dbName = "Regular Cheque.IMD"
			task.AddExtraction dbName, "", "(@Match(TRAN_TYPE,""BI"", ""CI"")  .AND.  INSTRMNT_TYPE == ""CHQ"")"
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Cheque Txns.IMD", "Get_Regular_Cheques", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("Regular Cheque.IMD") Then	
	Set db = Client.OpenDatabase("Regular Cheque.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CLG_ZONE_DATE_DATE"
		field.Description = ""
		field.Type = WI_VIRT_DATE
		field.Equation = "@CTOD(""00000000"", ""YYYYMMDD"")"
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Regular Cheque.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CLG_ZONE_CODE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """"""
		field.Length = 2
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

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

Function Get_Multiple_Cheques_deposited
' Get multiple cheques deposited by the same customer
If haveRecords("All Cheques.IMD") Then
	Set db = Client.OpenDatabase("All Cheques.IMD")
		Set task = db.DupKeyExclusion
			task.IncludeAllFields
			task.AddKey "CUST_ID", "A"
			task.DifferentField = "INSTRMNT_NUM"
			dbName = "Multiple_Cheques_deposited_to_same_account.IMD"
			task.CreateVirtualDatabase = False
			task.PerformTask dbName, ""
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("All Cheques.IMD", "Get_Multiple_Cheques_deposited", "DupKeyExclusion", "Error", "Databases empty or does not exist.")		
	
End If

	

If haveRecords("Multiple_Cheques_deposited_to_same_account.IMD") Then
	Set db = Client.OpenDatabase("Multiple_Cheques_deposited_to_same_account.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "COMP_CHQ_NUMBER"
				field.Description = ""
				field.Type = WI_VIRT_NUM
				field.Equation = "@VAL( INSTRMNT_NUM)"
				field.Decimals = 0
				task.AppendField field
				task.PerformTask
			Set field = Nothing
		Set task = Nothing
	Set db = Nothing
End If
End Function

Function Get_Sequential_Cheque_Deposited

If haveRecords("Multiple_Cheques_deposited_to_same_account.IMD") Then
	Set db = Client.OpenDatabase("Multiple_Cheques_deposited_to_same_account.IMD")
		Set task = db.Sort
			task.AddKey "ACID", "A"
			task.AddKey "TRAN_DATE_DATE", "A"
			task.AddKey "COMP_CHQ_NUMBER", "A"
			dbName = "Multiple_Cheques_deposited_to_same_account_Sort.IMD"
			task.PerformTask dbName
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Multiple_Cheques_deposited_to_same_account.IMD", "Get_Sequential_Cheque_Deposited", "Sort", "Error", "Databases empty or does not exist.")	
	
End If
	
If haveRecords("Multiple_Cheques_deposited_to_same_account_Sort.IMD") Then
	Set db = Client.OpenDatabase("Multiple_Cheques_deposited_to_same_account_Sort.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "SEQUENTIAL"
				field.Description = ""
				field.Type = WI_VIRT_NUM
'				field.Equation = "@If(COMP_CHQ_NUMBER== (@GetPreviousValue(""COMP_CHQ_NUMBER"") + 1) .OR. (COMP_CHQ_NUMBER+ 1) == @GetNextValue(""COMP_CHQ_NUMBER""), 1, 0)"
				field.Equation = "@If(COMP_CHQ_NUMBER== (@GetPreviousValue(""COMP_CHQ_NUMBER"") + 1)   .AND. @Abs(@Age(TRANSACTION_DATE,@GetPreviousValue(""TRANSACTION_DATE"")))<3 .OR. ((COMP_CHQ_NUMBER+ 1) == @GetNextValue(""COMP_CHQ_NUMBER"") .AND. @Abs(@Age(TRANSACTION_DATE,@GetNextValue(""TRANSACTION_DATE"")))<3) , 1, 0)"

				field.Decimals = 0
				task.AppendField field
				task.PerformTask
			Set field = Nothing
		Set task = Nothing
	Set db = Nothing	
End If
End Function

Function Get_Customer_Info
If haveRecords("Multiple_Cheques_deposited_to_same_account_Sort.IMD") And haveRecords("General Acct Master Lite.IMD")  Then
	Set db = Client.OpenDatabase("Multiple_Cheques_deposited_to_same_account_Sort.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "General Acct Master Lite.IMD"
			task.IncludeAllPFields
			task.AddSFieldToInc "FORACID"
			task.AddSFieldToInc "SCHM_DESC"
			task.AddSFieldToInc "CUST_NAME"
			task.AddSFieldToInc "CUST_PERM_ADDR1"
			task.AddSFieldToInc "CUST_PERM_ADDR2"
			task.AddSFieldToInc "CUSTOMER_TYPE"
			task.AddSFieldToInc "RISK_SCORE"
			task.AddSFieldToInc "RISK_CATEGORY"
			task.AddSFieldToInc "MONTHLY_DEPOSIT"
			task.AddSFieldToInc "ANNUAL_INCOME"
			task.AddMatchKey "ACID", "ACID", "A"
			task.CreateVirtualDatabase = False
			dbName = "Multiple_Cheques_deposited_to_same_account_Cust_Info.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Multiple_Cheques_deposited_to_same_account_Sort.IMD or General Acct Master Lite.IMD", "Get_Customer_Info", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If
End Function


' Data: Direct Extraction
Function Get_Multiple_Cheques_deposited_Details

If haveRecords("Multiple_Cheques_deposited_to_same_account_Cust_Info.IMD") Then
	Set db = Client.OpenDatabase("Multiple_Cheques_deposited_to_same_account_Cust_Info.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "CUST_ID"
	task.AddFieldToInc "CUST_NAME"
	task.AddFieldToInc "INSTRMNT_NUM"
	task.AddFieldToInc "TRAN_PARTICULAR"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMT"
	task.AddFieldToInc "CLG_ZONE_DATE_DATE"
	task.AddFieldToInc "CLG_ZONE_CODE"
	task.AddFieldToInc "TRAN_ID"
	task.AddFieldToInc "TRAN_TYPE"
	task.AddFieldToInc "TRANSACTION_CRNCY_CODE"
	task.AddFieldToInc "BR_NAME"
	task.AddFieldToInc "FORACID"
	task.AddFieldToInc "CUST_PERM_ADDR1"
	task.AddFieldToInc "CUST_PERM_ADDR2"
	task.AddFieldToInc "SCHM_DESC"
	task.AddFieldToInc "CUSTOMER_TYPE"
	task.AddFieldToInc "RISK_SCORE"
	task.AddFieldToInc "RISK_CATEGORY"
	task.AddFieldToInc "MONTHLY_DEPOSIT"
	task.AddFieldToInc "ANNUAL_INCOME"
	dbName = "AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD"
	task.AddExtraction dbName, "", "SEQUENTIAL == 1"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Multiple_Cheques_deposited_to_same_account_Cust_Info.IMD", "Get_Multiple_Cheques_deposited_Details", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If
	
If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 9
	task.ReplaceField "TRAN_ID", field
	
	Set field = db.TableDef.NewField
	field.Name = "CHEQUE_AMT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "TRANSACTION_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_TYPE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 1
	task.ReplaceField "TRAN_TYPE", field

	Set field = db.TableDef.NewField
	field.Name = "CURRENCY"
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
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
	
If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD")
	Set task = db.TableManagement
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
	field.Name = "ACCOUNT_DESC"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 25
	task.ReplaceField "SCHM_DESC", field
	
	Set field = db.TableDef.NewField
	field.Name = "CHEQUE_NUMBER"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "INSTRMNT_NUM", field
	
	Set field = db.TableDef.NewField
	field.Name = "CLEARING_ZONE_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "CLG_ZONE_DATE_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "CLEARING_ZONE_CODE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 10
	task.ReplaceField "CLG_ZONE_CODE", field
	
	Set field = db.TableDef.NewField
	field.Name = "RUN_DATE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Dtoc(@Date(),""YYYY-MM-DD"")"
	field.Length = 10
	task.AppendField field

	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function Get_Multiple_Cheques_deposited_Summ
		
If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "CUSTOMER_ID"
	task.AddFieldToInc "CUSTOMER_NAME"
	task.AddFieldToInc "CUSTOMER_ADDRESS1"
	task.AddFieldToInc "CUSTOMER_ADDRESS2"
	task.AddFieldToInc "CUSTOMER_TYPE"
	'task.AddFieldToInc "RISK_SCORE"
	'task.AddFieldToInc "RISK_CATEGORY"
	'task.AddFieldToInc "MONTHLY_DEPOSIT"
	'task.AddFieldToInc "ANNUAL_INCOME"
	task.AddFieldToTotal "CHEQUE_AMT"
	dbName = "AML31_Multiple_Cheques_deposited_to_same_Account_Summ_int.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.UseFieldFromFirstOccurrence = TRUE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	
Else
	Call logfile("AML31_Multiple_Cheques_deposited_to_same_Account_Details.IMD", "Get_Multiple_Cheques_deposited_Summ", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_int.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_int.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD"
	'task.AddExtraction dbName, "", "CHEQUE_AMT_SUM>@If(CUSTOMER_TYPE == ""Business"", " &  e_AML31_B_THRESH & ", " &  e_AML31_P_THRESH & ")  .AND. NO_OF_RECS >= @If(CUSTOMER_TYPE == ""Business""," &  e_AML31_B_COUNT & "," &  e_AML31_P_COUNT & ")"
	task.AddExtraction dbName, "",""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TOTAL_CHEQUE_AMT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "CHEQUE_AMT_SUM", field
	task.PerformTask
 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RUN_DATE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@Dtoc(@Date(),""YYYY-MM-DD"")"
	field.Length = 10
	task.AppendField field
	task.PerformTask
 
'	Set task = db.TableManagement
'	Set field = db.TableDef.NewField
'	field.Name = "NO_OF_SEQ_CHEQUES"
'	field.Description = "Number of records found for this key value"
'	field.Type = WI_NUM_FIELD
'	field.Equation = ""
'	field.Decimals = 0
'	task.ReplaceField "NO_OF_RECS", field
'	task.PerformTask
	
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If

If haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD")  And haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "CUSTOMER_ID", "CUSTOMER_ID", "A"
	task.CreateVirtualDatabase = False
	dbName = "AML31_Multiple_Cheques_deposited_to_same_Account_Details.IMD"
	task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
	
End Function

Function Join_RiskRating
If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_tmp.IMD")	
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
		dbName = "AML31_Multiple_Cheques_deposited_to_same_Account_Summ.IMD"
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
	Set task = Nothing
	Set db = Nothing
End If
	
End Function

Function CleanUp
	DeleteFile("Foreign Cheque Details.IMD")
	DeleteFile("Foreign Cheques.IMD")
	DeleteFile("Clearing Cheques.IMD")
	DeleteFile("Regular Cheque.IMD")
	DeleteFile("Cheques_Deposited.IMD")
	DeleteFile("All Cheques.IMD")
	'DeleteFile("Cheque Txns.IMD")
	DeleteFile("Clearing Cheques.IMD")
	DeleteFile( "Regular Cheque.IMD")
	DeleteFile("Active Regularised OCI.IMD")
	DeleteFile("HTD OCP.IMD")  
	DeleteFile("Active Regularised OCP.IMD")
	DeleteFile("Multiple_Cheques_deposited_to_same_account.IMD")
	DeleteFile("Multiple_Cheques_deposited_to_same_account_Sort.IMD")
	DeleteFile("Multiple_Cheques_deposited_to_same_account_Cust_Info.IMD")
	DeleteFile("Multiple_Cheques_deposited_to_same_account_Acct_Info")
	DeleteFile("AML31_Multiple_Cheques_deposited_to_same_Account_Details.IMD")
	'DeleteFile("AML31_Multiple_Cheques_deposited_to_same_Account_Summ.IMD") 
	DeleteFile("AML31_Multiple_Cheques_deposited_to_same_Account_Summ_int.IMD")
	DeleteFile("AML31_Multiple_Cheques_deposited_to_same_Account_Details_int.IMD")
	DeleteFile("FEX ACCT ENTRY EXTRACT.IMD")
	DeleteFile("Instrument Acct Entry.IMD")
	DeleteFile("Instrument Acct Branch.IMD")
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

'Check if the named database contains records
'--------------------------------------------------------------------------------
Function ContainsRecords(ByVal dbName As String) As Boolean
'--------------------------------------------------------------------------------
	Dim records As Double
	Dim db As Object
	Dim rs As Object
	
	records = 0
	ContainsRecords = False
	If fso.FileExists(dbName) Then
		Set db = Client.OpenDatabase(dbName)
			Set rs =  db.RecordSet
				If rs.count > 0 Then
					ContainsRecords = True
				Else
					errors_string = errors_string & " with errors -" & dbname & " has no records." & Chr(10)
				End If
			Set rs = Nothing
		db.close
		Set db = Nothing
	Else
		errors_string = errors_string & " with errors -" & dbname & " missing." & Chr(10)
		
	End If
End Function




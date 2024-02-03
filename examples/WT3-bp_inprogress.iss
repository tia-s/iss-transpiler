'====================================================================================================
'	Test#: 		WT3 - Offshore Wire after Incoming Wire
'	Risk:		Accounts may be used for facilitating off-shore banking.
' 	Objective:	Identify wire transfers to off-shore banks after an incoming wire transfer.
' 	Frequency:	Daily
' 	Last Modified:	07/27/2016
'	Comments:	Seperate Incoming and Outgoing details to facilitate grouped display
'====================================================================================================
'	Script Dependencies: Daily Interim Script
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "WT3 - Offshore Wire after Incoming Wire"
Const scriptname_log ="WT3 - Offshore Wire after Incoming Wire.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	Client.closeAll
	CleanUp
	
	Join_FX_Bills_Master_GAM
	Get_ITT_OTT_Daily
	Join_ITT_OTT_All_Matches
	Get_Outgoing_After_Incoming_Wire
	Sum_Wires
	ChangeField
	Sum_Final
	Rename
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


	Client.CloseAll
	CleanUp
	'Client.Quit
End Sub

' Collect account information for wires.
Function Join_FX_Bills_Master_GAM
	If haveRecords("TBAADM.FX_BILL_MASTER_TABLE.IMD") Then
	Set db = Client.OpenDatabase("TBAADM.FX_BILL_MASTER_TABLE.IMD")
		Set task = db.Extraction
			task.AddFieldToInc  "BILL_ID"
			task.AddFieldToInc  "OPER_ACID"
			task.AddFieldToInc  "ENTITY_CRE_FLG"
			task.AddFieldToInc  "DEL_FLG"
			task.AddFieldToInc  "REG_TYPE"
			task.AddFieldToInc  "REG_SUB_TYPE"
			task.AddFieldToInc  "BILL_DATE_DATE"
			task.AddFieldToInc  "BILL_DATE_TIME"
			task.AddFieldToInc  "BILL_CNTRY_CODE"
			task.AddFieldToInc  "BILL_CRNCY_CODE"
			task.AddFieldToInc  "BILL_AMT"
			task.AddFieldToInc  "BILL_AMT_INR"
			task.AddFieldToInc  "PARTY_CODE"
			task.AddFieldToInc  "PARTY_NAME"
			task.AddFieldToInc  "PARTY_ADDR1"
			task.AddFieldToInc  "PARTY_ADDR2"
			task.AddFieldToInc  "PARTY_ADDR3"
			task.AddFieldToInc  "PARTY_CNTRY_CODE"
			task.AddFieldToInc  "OTHER_PARTY_CNTRY_CODE"
			task.AddFieldToInc  "OTHER_PARTY_NAME"
			task.AddFieldToInc  "OTHER_PARTY_ADDR_1"
			task.AddFieldToInc  "OTHER_PARTY_ADDR_2"
			task.AddFieldToInc  "OTHER_PARTY_ADDR_3"
			task.AddFieldToInc  "OTHER_PARTY_CODE"
			task.AddFieldToInc  "RCRE_USER_ID"
			task.AddFieldToInc  "RCRE_TIME_DATE"
			task.AddFieldToInc  "RCRE_TIME_TIME"
		dbName = "FX_Bill_Extract.IMD"
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("TBAADM.FX_BILL_MASTER_TABLE.IMD", "Join_FX_Bills_Master_GAM", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("FX_Bill_Extract.IMD") And haveRecords("General Acct Master Lite.IMD") Then

	Set db = Client.OpenDatabase("FX_Bill_Extract.IMD")
		Set task = db.VisualConnector
			id0 = task.AddDatabase ("FX_Bill_Extract.IMD")
			id1 = task.AddDatabase ("General Acct Master Lite.IMD")
			task.MasterDatabase = id0
			task.AppendDatabaseNames = FALSE
			task.IncludeAllPrimaryRecords = FALSE
			task.AddRelation id0, "OPER_ACID", id1, "ACID"
			task.AddFieldToInclude id0, "BILL_ID"
			task.AddFieldToInclude id0, "ENTITY_CRE_FLG"
			task.AddFieldToInclude id0, "DEL_FLG"
			task.AddFieldToInclude id0, "REG_TYPE"
			task.AddFieldToInclude id0, "REG_SUB_TYPE"
			task.AddFieldToInclude id0, "BILL_DATE_DATE"
			task.AddFieldToInclude id0, "BILL_DATE_TIME"
			task.AddFieldToInclude id0, "BILL_CNTRY_CODE"
			task.AddFieldToInclude id0, "BILL_CRNCY_CODE"
			task.AddFieldToInclude id0, "BILL_AMT"
			task.AddFieldToInclude id0, "BILL_AMT_INR"
			task.AddFieldToInclude id0, "PARTY_CODE"
			task.AddFieldToInclude id0, "PARTY_NAME"
			task.AddFieldToInclude id0, "PARTY_ADDR1"
			task.AddFieldToInclude id0, "PARTY_ADDR2"
			task.AddFieldToInclude id0, "PARTY_ADDR3"
			task.AddFieldToInclude id0, "PARTY_CNTRY_CODE"
			task.AddFieldToInclude id0, "OTHER_PARTY_CNTRY_CODE"
			task.AddFieldToInclude id0, "OTHER_PARTY_NAME"
			task.AddFieldToInclude id0, "OTHER_PARTY_ADDR_1"
			task.AddFieldToInclude id0, "OTHER_PARTY_ADDR_2"
			task.AddFieldToInclude id0, "OTHER_PARTY_ADDR_3"
			task.AddFieldToInclude id0, "OTHER_PARTY_CODE"
			task.AddFieldToInclude id0, "RCRE_USER_ID"
			task.AddFieldToInclude id0, "RCRE_TIME_DATE"
			task.AddFieldToInclude id0, "RCRE_TIME_TIME"
			task.AddFieldToInclude id1, "CUST_ID"
			task.AddFieldToInclude id1, "FORACID"
			task.AddFieldToInclude id1, "BACID"
			task.AddFieldToInclude id1, "ACCT_OWNERSHIP"
			task.AddFieldToInclude id1, "ACCT_NAME"
			task.AddFieldToInclude id1, "SCHM_DESC"
			task.CreateVirtualDatabase = False
			dbName = "FX Bill GAM.IMD"
			task.OutputDatabaseName = dbName
			task.PerformTask
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("FX_Bill_Extract.IMD or General Acct Master Lite.IMD", "Join_FX_Bills_Master_GAM", "VisualConnector", "Error", "Databases empty or does not exist.")	

End If
End Function

Function Get_ITT_OTT_Daily
'exclude internal transaction
If haveRecords("FX Bill GAM.IMD") Then
	Set db = Client.OpenDatabase("FX Bill GAM.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_CUST_ID"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@IF(@ALLTRIM(CUST_ID) <> """", CUST_ID, PARTY_CODE)"
		field.Length = 9
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("FX Bill GAM.IMD") Then
	Set db = Client.OpenDatabase("FX Bill GAM.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "FXBills_ITT_Daily.IMD"
		task.AddExtraction dbName, "", "REG_TYPE =""ITT"" .and. DEL_FLG <> ""Y"" .AND. BACID <> ""3133170000"" .AND. COMP_CUST_ID<> """""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("FX Bill GAM.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "FXBills_OTT_Daily.IMD"
		task.AddExtraction dbName, "", "REG_TYPE =""OTT"" .AND. DEL_FLG <> ""Y"" .AND. BACID <> ""3133170000"" .AND. COMP_CUST_ID<> """""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End if
	
If haveRecords("FXBills_OTT_Daily.IMD") Then
'Get all customer ids present
	Set db = Client.OpenDatabase("FXBills_OTT_Daily.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "CUST_ID1"
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

If haveRecords("FXBills_OTT_Daily.IMD") Then
	Set db = Client.OpenDatabase("FXBills_OTT_Daily.IMD")
		Set task = db.TableManagement
			Set field = db.TableDef.NewField
				field.Name = "CUST_ID"
				field.Description = ""
				field.Type = WI_CHAR_FIELD
				field.Equation = ""
				field.Length = 9
				task.ReplaceField "COMP_CUST_ID", field
				task.PerformTask
			Set field = Nothing
		Set task = Nothing
	Set db = Nothing	
End If
	
If haveRecords("FXBills_ITT_Daily.IMD") Then
'Get all customer ids present
	Set db = Client.OpenDatabase("FXBills_ITT_Daily.IMD")
	Set task = db.TableManagement
		Set field = db.TableDef.NewField
			field.Name = "CUST_ID1"
			field.Description = ""
			field.Type = WI_CHAR_FIELD
			field.Equation = ""
			field.Length = 9
			task.ReplaceField "CUST_ID", field
			task.PerformTask
End If

If haveRecords("FXBills_ITT_Daily.IMD") Then
	Set db = Client.OpenDatabase("FXBills_ITT_Daily.IMD")
	Set task = db.TableManagement
		Set field = db.TableDef.NewField
			field.Name = "CUST_ID"
			field.Description = ""
			field.Type = WI_CHAR_FIELD
			field.Equation = ""
			field.Length = 9
			task.ReplaceField "COMP_CUST_ID", field
			task.PerformTask
		Set field = Nothing
	Set task = Nothing
	Set db = Nothing	
End If
End Function

Function Join_ITT_OTT_All_Matches
If haveRecords("FXBills_OTT_Daily.IMD") Then
	Set db = Client.OpenDatabase("FXBills_OTT_Daily.IMD")
	Set task = db.VisualConnector
		id0 = task.AddDatabase ("FXBills_OTT_Daily.IMD")
		id1 = task.AddDatabase ("FXBills_ITT_Daily.IMD")
		task.MasterDatabase = id0
		task.AppendDatabaseNames = FALSE
		task.IncludeAllPrimaryRecords = TRUE
		task.AddRelation id0, "CUST_ID", id1, "CUST_ID"
		task.IncludeAllFields
		task.CreateVirtualDatabase = False
		dbName = "FXBills_Daily_ITTwOTT.IMD"
		task.OutputDatabaseName = dbName
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("FXBills_OTT_Daily.IMD or FXBills_ITT_Daily.IMD", "Join_ITT_OTT_All_Matches", "VisualConnector", "Error", "Databases empty or does not exist.")	

End If
End Function


Function Get_Outgoing_After_Incoming_Wire
' Extract Detail Records that are wires outgoing after incoming (comparison of timestamp)

If haveRecords("FXBills_Daily_ITTwOTT.IMD") Then
	Set db = Client.OpenDatabase("FXBills_Daily_ITTwOTT.IMD")
		Set task = db.Extraction
			task.AddFieldToInc "BILL_ID"
			task.AddFieldToInc "BILL_ID1"
			task.AddFieldToInc "BILL_DATE_DATE1"
			task.AddFieldToInc "BILL_CRNCY_CODE1"
			task.AddFieldToInc "BILL_AMT1"
			task.AddFieldToInc "BILL_AMT_INR1"
			task.AddFieldToInc "PARTY_CODE1"
			task.AddFieldToInc "PARTY_NAME1"
			task.AddFieldToInc "PARTY_ADDR11"
			task.AddFieldToInc "OTHER_PARTY_CODE1"
			task.AddFieldToInc "OTHER_PARTY_NAME1"
			task.AddFieldToInc "OTHER_PARTY_ADDR_11"
			task.AddFieldToInc "OTHER_PARTY_CNTRY_CODE1"
			task.AddFieldToInc "RCRE_TIME_DATE1"
			task.AddFieldToInc "RCRE_TIME_TIME1"
			task.AddFieldToInc "CUST_ID"
			task.AddFieldToInc "FORACID1"
			task.AddFieldToInc "BACID"
			task.AddFieldToInc "ACCT_NAME1"
			task.AddFieldToInc "SCHM_DESC1"
			dbName = "Outgoing After Incoming Details.IMD"
			task.AddExtraction dbName, "", "BILL_ID1 <> """"  .AND. (RCRE_TIME_TIME > RCRE_TIME_TIME1) .AND. (RCRE_TIME_DATE = RCRE_TIME_DATE1) .AND. OTHER_PARTY_CNTRY_CODE <> ""JM"""
			task.CreateVirtualDatabase = False
			task.PerformTask 1, db.Count
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("FXBills_Daily_ITTwOTT.IMD", "Get_Outgoing_After_Incoming_Wire", "Direct Extraction", "Error", "Databases empty or does not exist.")		
	
End If

' Collect outgoing wire details
If haveRecords("FXBills_OTT_Daily.IMD") And haveRecords("Outgoing After Incoming Details.IMD") Then
	Set db = Client.OpenDatabase("FXBills_OTT_Daily.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "Outgoing After Incoming Details.IMD"
			task.AddPFieldToInc "BILL_ID"
			task.AddPFieldToInc "BILL_DATE_DATE"
			task.AddPFieldToInc "BILL_CRNCY_CODE"
			task.AddPFieldToInc "BILL_AMT"
			task.AddPFieldToInc "BILL_AMT_INR"
			task.AddPFieldToInc "PARTY_CODE"
			task.AddPFieldToInc "PARTY_NAME"
			task.AddPFieldToInc "PARTY_ADDR1"
			task.AddPFieldToInc "OTHER_PARTY_NAME"
			task.AddPFieldToInc "OTHER_PARTY_ADDR_1"
			task.AddPFieldToInc "OTHER_PARTY_CNTRY_CODE"
			task.AddPFieldToInc "RCRE_TIME_DATE"
			task.AddPFieldToInc "RCRE_TIME_TIME"
			task.AddPFieldToInc "CUST_ID"
			task.AddPFieldToInc "FORACID"
			task.AddPFieldToInc "ACCT_NAME"
			task.AddPFieldToInc "SCHM_DESC"
			task.AddPFieldToInc "BACID"
			task.AddSFieldToInc "BILL_ID"			
			task.AddMatchKey "BILL_ID", "BILL_ID", "A"
			task.CreateVirtualDatabase = False
			dbName = "Outgoing Wire Details.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("FXBills_OTT_Daily.IMD or Outgoing After Incoming Details.IMD", "Get_Outgoing_After_Incoming_Wire", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If

' Collect incoming wire details
If haveRecords("FXBills_ITT_Daily.IMD") And haveRecords("Outgoing After Incoming Details.IMD") Then
	Set db = Client.OpenDatabase("FXBills_ITT_Daily.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "Outgoing After Incoming Details.IMD"
			task.AddPFieldToInc "BILL_ID"
			task.AddPFieldToInc "BILL_DATE_DATE"
			task.AddPFieldToInc "BILL_CRNCY_CODE"
			task.AddPFieldToInc "BILL_AMT"
			task.AddPFieldToInc "BILL_AMT_INR"
			task.AddPFieldToInc "PARTY_CODE"
			task.AddPFieldToInc "PARTY_NAME"
			task.AddPFieldToInc "PARTY_ADDR1"
			task.AddPFieldToInc "OTHER_PARTY_NAME"
			task.AddPFieldToInc "OTHER_PARTY_ADDR_1"
			task.AddPFieldToInc "OTHER_PARTY_CNTRY_CODE"
			task.AddPFieldToInc "RCRE_TIME_DATE"
			task.AddPFieldToInc "RCRE_TIME_TIME"
			task.AddPFieldToInc "CUST_ID"
			task.AddPFieldToInc "FORACID"
			task.AddPFieldToInc "ACCT_NAME"
			task.AddPFieldToInc "SCHM_DESC"
			task.AddPFieldToInc "BACID"
			task.AddSFieldToInc "BILL_ID1"			
			task.AddMatchKey "BILL_ID", "BILL_ID1", "A"
			task.CreateVirtualDatabase = False
			dbName = "Incoming Wire Details.IMD"
			task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
Else
	Call logfile("FXBills_ITT_Daily.IMD or Outgoing After Incoming Details.IMD", "Get_Outgoing_After_Incoming_Wire", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If
End Function


' Count the number of incoming and outgoing wires
Function Sum_Wires

If haveRecords("Outgoing Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Outgoing Wire Details.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "CUST_ID"
		task.AddFieldToInc "PARTY_CODE"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDR1"
		task.AddFieldToInc "RCRE_TIME_DATE"
		task.AddFieldToTotal "BILL_AMT_INR"
		dbName = "Outgoing_Sum.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Outgoing Wire Details.IMD", "Sum_Wires", "Summarization", "Error", "Databases empty or does not exist.")		
	
End If

If haveRecords("Incoming Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Incoming Wire Details.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "CUST_ID"
		task.AddFieldToInc "PARTY_CODE"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDR1"
		task.AddFieldToInc "RCRE_TIME_DATE"
		task.AddFieldToTotal "BILL_AMT_INR"
		dbName = "Incoming_Sum.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.UseFieldFromFirstOccurrence = TRUE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Incoming Wire Details.IMD", "Sum_Wires", "Summarization", "Error", "Databases empty or does not exist.")		
	
End If	
End Function

' Rename fields
Function ChangeField
If haveRecords("Outgoing_Sum.IMD") Then
	Set db = Client.OpenDatabase("Outgoing_Sum.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_OUTGOING_WIRES"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	task.PerformTask

	Set field = db.TableDef.NewField
	field.Name = "TOTAL_AMT_OUTGOING_JM"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT_INR_SUM", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Incoming_Sum.IMD") Then
	Set db = Client.OpenDatabase("Incoming_Sum.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NO_OF_INCOMING_WIRES"
	field.Description = "Number of records found for this key value"
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 0
	task.ReplaceField "NO_OF_RECS", field
	task.PerformTask

	Set field = db.TableDef.NewField
	field.Name = "TOTAL_AMT_INCOMING_JM"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT_INR_SUM", field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function Sum_Final
If haveRecords("Outgoing_Sum.IMD") And haveRecords("Incoming_Sum.IMD") Then
	Set db = Client.OpenDatabase("Outgoing_Sum.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Incoming_Sum.IMD"
		task.AddPFieldToInc "CUST_ID"
		task.AddPFieldToInc "PARTY_NAME"
		task.AddPFieldToInc "PARTY_ADDR1"
		task.AddPFieldToInc "RCRE_TIME_DATE"
		task.AddSFieldToInc "NO_OF_INCOMING_WIRES"
		task.AddSFieldToInc "TOTAL_AMT_INCOMING_JM"
		task.AddPFieldToInc "NO_OF_OUTGOING_WIRES"
		task.AddPFieldToInc "TOTAL_AMT_OUTGOING_JM"
		task.AddMatchKey "CUST_ID", "CUST_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Incoming_Outgoing_wires_Summ_tmp.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
Else
	Call logfile("Outgoing_Sum.IMD or Incoming_Sum.IMD", "Sum_Final", "JoinDatabase", "Error", "Databases empty or does not exist.")	

End If
End Function


	Function Rename
'Rename Details
If haveRecords("Incoming Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Incoming Wire Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "INCOMING_BILL_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 15
	task.ReplaceField "BILL_ID1", field
	
	Set field = db.TableDef.NewField
	field.Name = "BILL_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "BILL_DATE_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "CURRENCY"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "BILL_CRNCY_CODE", field

	Set field = db.TableDef.NewField
	field.Name = "INCOMING_WIRE_AMT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "INCOMING_WIRE_AMT_JMD"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT_INR", field

	Set field = db.TableDef.NewField
	field.Name = "INCOMING_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "RCRE_TIME_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "INCOMING_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = ""
	task.ReplaceField "RCRE_TIME_TIME", field

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
	field.Name = "ACCOUNT_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "ACCT_NAME", field

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

If haveRecords("Outgoing Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Outgoing Wire Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "OUTGOING_BILL_ID"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 15
	task.ReplaceField "BILL_ID1", field
	
	Set field = db.TableDef.NewField
	field.Name = "BILL_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "BILL_DATE_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "CURRENCY"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "BILL_CRNCY_CODE", field

	Set field = db.TableDef.NewField
	field.Name = "OUTGOING_WIRE_AMT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT", field

	Set field = db.TableDef.NewField
	field.Name = "OUTGOING_WIRE_AMT_JMD"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 4
	task.ReplaceField "BILL_AMT_INR", field

	Set field = db.TableDef.NewField
	field.Name = "OUTGOING_DATE"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "RCRE_TIME_DATE", field

	Set field = db.TableDef.NewField
	field.Name = "OUTGOING_TIME"
	field.Description = ""
	field.Type = WI_TIME_FIELD
	field.Equation = ""
	task.ReplaceField "RCRE_TIME_TIME", field

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
	field.Name = "ACCOUNT_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 80
	task.ReplaceField "ACCT_NAME", field

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

Function Join_RiskRating
'If haveRecords("ACID_FORACID.IMD") And haveRecords("Incoming_Outgoing_wires_Summ_tmp.IMD") Then
'	Set db = Client.OpenDatabase("Incoming_Outgoing_wires_Summ_tmp.IMD")	
'	Set task = db.JoinDatabase
'		task.FileToJoin "ACID_FORACID.IMD" 
'		task.IncludeAllPFields
'		task.AddSFieldToInc "ACID"
'		task.AddMatchKey "ACCOUNT_NUMBER", "FORACID", "A"
'		task.CreateVirtualDatabase = False
'		dbName = "Incoming_Outgoing_wires_Summ_tmp1.IMD"
'		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
'	Set task = Nothing
'	Set db = Nothing
'End If


If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("Incoming_Outgoing_wires_Summ_tmp.IMD") Then
	Set db = Client.OpenDatabase("Incoming_Outgoing_wires_Summ_tmp.IMD")	
	Set task = db.JoinDatabase
		task.FileToJoin "Customer_Turnover_wRisk.IMD" 
		task.IncludeAllPFields
		task.AddSFieldToInc "MONTHLY_DEPOSIT"
		task.AddSFieldToInc "RISK_SCORE"
		task.AddSFieldToInc "RISK_CATEGORY"
		task.AddSFieldToInc "OCCUPATION"
		task.AddSFieldToInc "INDUSTRY"
		task.AddSFieldToInc "ANNUAL_INCOME"
		task.AddMatchKey "CUST_ID", "CUST_ID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Incoming_Outgoing_wires_Summ.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExportDatabase

If haveRecords("Outgoing Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Outgoing Wire Details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "BILL_DATE"
		task.AddFieldToInc "CURRENCY"
		task.AddFieldToInc "OUTGOING_WIRE_AMT"
		task.AddFieldToInc "OUTGOING_WIRE_AMT_JMD"
		task.AddFieldToInc "PARTY_CODE"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDR1"
		task.AddFieldToInc "OTHER_PARTY_NAME"
		task.AddFieldToInc "OTHER_PARTY_ADDR_1"
		task.AddFieldToInc "OTHER_PARTY_CNTRY_CODE"
		task.AddFieldToInc "OUTGOING_DATE"
		task.AddFieldToInc "OUTGOING_TIME"
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "OUTGOING_BILL_ID"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\WT3 Outgoing Wire Details.MDB", "Database", "MDB2000", 1, db.Count, eqn
 
	Set db = Nothing
	Set task = Nothing
End If

If haveRecords("Incoming Wire Details.IMD") Then
	Set db = Client.OpenDatabase("Incoming Wire Details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "BILL_DATE"
		task.AddFieldToInc "CURRENCY"
		task.AddFieldToInc "INCOMING_WIRE_AMT"
		task.AddFieldToInc "INCOMING_WIRE_AMT_JMD"
		task.AddFieldToInc "PARTY_CODE"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDR1"
		task.AddFieldToInc "OTHER_PARTY_NAME"
		task.AddFieldToInc "OTHER_PARTY_ADDR_1"
		task.AddFieldToInc "OTHER_PARTY_CNTRY_CODE"
		task.AddFieldToInc "INCOMING_DATE"
		task.AddFieldToInc "INCOMING_TIME"
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "ACCOUNT_DESCRIPTION"
		task.AddFieldToInc "INCOMING_BILL_ID"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\WT3 Incoming Wire Details.MDB", "Database", "MDB2000", 1, db.Count, eqn
	 
	Set db = Nothing
	Set task = Nothing
End If


If haveRecords("Incoming_Outgoing_wires_Summ.IMD") Then
	Set db = Client.OpenDatabase("Incoming_Outgoing_wires_Summ.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUST_ID"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDR1"
		task.AddFieldToInc "RCRE_TIME_DATE"
		task.AddFieldToInc "NO_OF_INCOMING_WIRES"
		task.AddFieldToInc "TOTAL_AMT_INCOMING_JM"
		task.AddFieldToInc "NO_OF_OUTGOING_WIRES"
		task.AddFieldToInc "TOTAL_AMT_OUTGOING_JM"
		'task.AddFieldToInc "CUST_TRN"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\WT3 Incoming then Outgoing Wire Summary.MDB", "Database", "MDB2000", 1, db.Count, eqn
RESULTSLOG(db.name)
			Set task = Nothing
	Set db = Nothing
	
	Else 
NORESULTSLOG("WT3_SUMMARY.IMD") 
	
End If
	
End Function

Function CleanUp
	DeleteFile("FX Bill GAM.IMD")
	DeleteFile("Outgoing_Sum.IMD")
	DeleteFile("Incoming_Sum.IMD")
	DeleteFile("FXBills_ITT_Daily.IMD")
	DeleteFile("FXBills_OTT_Daily.IMD")
	DeleteFile("FXBills_Daily_ITTwOTT.IMD")
	DeleteFile("Outgoing Wire Details.IMD")
	DeleteFile("Incoming Wire Details.IMD")
	DeleteFile("Outgoing After Incoming Summary.IMD")
	DeleteFile("Outgoing After Incoming Details.IMD")
	'DeleteFile("Incoming_Outgoing_wires_Summ.IMD")
	DeleteFile("Incoming_Outgoing_wires_Summ_tmp.IMD")
	DeleteFile("FX_Bill_Extract.IMD")
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

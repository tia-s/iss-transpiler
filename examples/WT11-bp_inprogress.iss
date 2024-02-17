'====================================================================================================
'	Test#: 		WT11 - Wires from Multiple Senders Over Threshold
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	18-Mar-14 3:20:45 AM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----
Const RESULT_FILENAME = "WT11 - Wires from Multiple Senders Over Threshold"
Const scriptname_log ="WT11 - Wires from Multiple Senders Over Threshold.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object

Sub Main

	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine

	'	Client.CloseAll
	CleanUp
	
	Call SummarizeInWireOthPrty()
	Call SummarizeInWireParty()
	Call GetMultiSendWiresOvrThrsh()
	Call GetMultiSendWiresDet()
	Call ModifyFieldForExport()
	Call Join_RiskRating()
	Call ExportSummDetails()
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



'Total wires in by other party and party.
Function SummarizeInWireOthPrty
If haveRecords("Inward Wire Cust Acct Branch.IMD") Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "USD"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """USD"""
		field.Length = 3
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
		
	Set field = db.TableDef.NewField
		field.Name = "RUN_DATE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@Dtoc(@Date(),""YYYY-MM-DD"")"
		field.Length = 10
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
	task.AppendField field
	
	task.PerformTask
	Set field = Nothing	
	Set task = Nothing
	Set db = Nothing
End If 	

If haveRecords("Inward Wire Cust Acct Branch.IMD") Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Current_Rate.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "VAR_CRNCY_UNITS"
		task.AddMatchKey "USD", "FXD_CRNCY_CODE", "A"
		task.CreateVirtualDatabase = False
	dbName = "Inward Wire Cust Acct Branch_US.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing	
End If

If haveRecords("Inward Wire Cust Acct Branch_US.IMD") Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch_US.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "COMP_USD_EQUIVALENT"
		field.Description = ""
		field.Type = WI_VIRT_NUM
		field.Equation = "@IF(BILL_CRNCY_CODE == ""USD"",  BILL_AMT, BILL_AMT_INR/ VAR_CRNCY_UNITS)"
		field.Decimals = 4
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Inward Wire Cust Acct Branch_US.IMD") Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch_US.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "CUST_ID"
		task.AddFieldToSummarize "PARTY_NAME"
		task.AddFieldToSummarize "PARTY_ADDR1"
		task.AddFieldToSummarize "OTHER_PARTY_NAME"
		task.AddFieldToSummarize "OTHER_PARTY_ADDR_1"
	'	task.AddFieldToSummarize "BILL_DATE_DATE"
	'	task.AddFieldToSummarize "DATE_OF_REMIT_DATE"
		task.AddFieldToTotal "COMP_USD_EQUIVALENT"
		task.AddFieldToTotal "BILL_AMT_INR"
		task.AddFieldToInc "RUN_DATE"
		dbName = "SummInWireOtherParty.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Analysis: Summarization
Function SummarizeInWireParty
If haveRecords("SummInWireOtherParty.IMD") Then
	Set db = Client.OpenDatabase("SummInWireOtherParty.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "CUST_ID"
		task.AddFieldToSummarize "PARTY_NAME"
		task.AddFieldToSummarize "PARTY_ADDR1"
	'	task.AddFieldToSummarize "BILL_DATE_DATE"
	'	task.AddFieldToSummarize "DATE_OF_REMIT_DATE"
		task.AddFieldToTotal "NO_OF_RECS"
		task.AddFieldToTotal "BILL_AMT_INR_SUM"
		task.AddFieldToTotal "COMP_USD_EQUIVALENT_SUM"
		task.AddFieldToInc "RUN_DATE"
		dbName = "SummInWireParty.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Data: Direct Extraction
Function GetMultiSendWiresOvrThrsh
If haveRecords("SummInWireParty.IMD") Then
	Set db = Client.OpenDatabase("SummInWireParty.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "Wires From Multiple Senders Over Threshold_tmp.IMD"
		'task.AddExtraction dbName, "", "( NO_OF_RECS1  >= " & e_Minimum_Wires_Count & " .AND.  COMP_USD_EQUIVALENT_SUM_SUM  >= "& e_TTR_Value_US &")"
		task.AddExtraction dbName, "", "NO_OF_RECS1  >= 3 .AND. COMP_USD_EQUIVALENT_SUM_SUM  >= 15000"
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End if
End Function

' File: Join Databases
Function GetMultiSendWiresDet
If haveRecords("Inward Wire Cust Acct Branch.IMD") And haveRecords("Wires From Multiple Senders Over Threshold_tmp.IMD")  Then
	Set db = Client.OpenDatabase("Inward Wire Cust Acct Branch.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "Wires From Multiple Senders Over Threshold_tmp.IMD"
		task.IncludeAllPFields
		task.IncludeAllSFields
		task.AddMatchKey "PARTY_NAME", "PARTY_NAME", "A"
		task.CreateVirtualDatabase = False
		dbName = "Wires from Multiple Senders Over Thrsh - Details.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End if
End Function

' Modify Fields in the details file for export
Function ModifyFieldForExport
If haveRecords("Wires From Multiple Senders Over Threshold_tmp.IMD") Then
	Set db = Client.OpenDatabase("Wires From Multiple Senders Over Threshold_tmp.IMD")
	Set task = db.TableManagement
	
	Set field = db.TableDef.NewField
		field.Name = "PARTY_ADDRESS_1"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 45
		task.ReplaceField "PARTY_ADDR1", field

	Set field = db.TableDef.NewField
		field.Name = "NUMBER_OF_WIRES"
		field.Description = "Number of records found for this key value"
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 0
		task.ReplaceField "NO_OF_RECS_SUM", field

	Set field = db.TableDef.NewField
		field.Name = "TOTAL_BILL_AMOUNT_JMD"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT_INR_SUM_SUM", field
				
	task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
	

End If

' Details
If haveRecords("Wires from Multiple Senders Over Thrsh - Details.IMD") Then
	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "BILL_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "BILL_DATE_DATE", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "BILL_CURRENCY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "BILL_CRNCY_CODE", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "BILL_AMOUNT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "BILL_AMOUNT_JMD"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT_INR", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "PARTY_ADDRESS_1"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 45
		task.ReplaceField "PARTY_ADDR1", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "PARTY_ADDRESS_2"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 45
		task.ReplaceField "PARTY_ADDR2", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "OTHER_PARTY_ADDRESS_1"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 45
		task.ReplaceField "OTHER_PARTY_ADDR_1", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "OTHER_PARTY_ADDRESS_2"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 45
		task.ReplaceField "OTHER_PARTY_ADDR_2", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "ACCOUNT_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "FORACID", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing

	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "CUST_NAME", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
	
	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "DATE_OF_REMIT"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "DATE_OF_REMIT_DATE", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
	
End If
End Function

Function Join_RiskRating
If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("Wires From Multiple Senders Over Threshold_tmp.IMD") Then
	Set db = Client.OpenDatabase("Wires From Multiple Senders Over Threshold_tmp.IMD")	
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
		dbName = "Wires From Multiple Senders Over Threshold.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function
' File - Export Database: MDB2000
Function ExportSummDetails
If haveRecords("Wires From Multiple Senders Over Threshold.IMD") Then
	Set db = Client.OpenDatabase("Wires From Multiple Senders Over Threshold.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDRESS_1"
'		task.AddFieldToInc "BILL_DATE"
'		task.AddFieldToInc "DATE_OF_REMIT"
		task.AddFieldToInc "NUMBER_OF_WIRES"
	'	task.AddFieldToInc "TOTAL_BILL_AMOUNT"
		task.AddFieldToInc "TOTAL_BILL_AMOUNT_JMD"
		task.AddFieldToInc "RUN_DATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory &  "Reports\WT11_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
		RESULTSLOG(db.name)
	Set task = Nothing
	Set db = Nothing

Else 
NORESULTSLOG("WT11_SUMMARY.IMD") 
End If

If haveRecords("Wires from Multiple Senders Over Thrsh - Details.IMD") Then
	Set db = Client.OpenDatabase("Wires from Multiple Senders Over Thrsh - Details.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "ACCOUNT_ID"
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "BILL_ID"
		task.AddFieldToInc "SOL_ID"
		task.AddFieldToInc "BILL_DATE"
		task.AddFieldToInc "DATE_OF_REMIT"
		task.AddFieldToInc "BILL_CURRENCY"
		task.AddFieldToInc "BILL_AMOUNT"
		task.AddFieldToInc "BILL_AMOUNT_JMD"
		task.AddFieldToInc "PARTY_NAME"
		task.AddFieldToInc "PARTY_ADDRESS_1"
		task.AddFieldToInc "PARTY_ADDRESS_2"
		task.AddFieldToInc "OTHER_PARTY_NAME"
		task.AddFieldToInc "OTHER_PARTY_ADDRESS_1"
		task.AddFieldToInc "OTHER_PARTY_ADDRESS_2"
		task.AddFieldToInc "RUN_DATE"
		eqn = ""
		task.PerformTask Client.WorkingDirectory &  "Reports\WT11_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function CleanUp
	DeleteFile("SummInWireOtherParty.IMD") 
	DeleteFile("SummInWireParty.IMD") 
	DeleteFile("SummInWirePartyFXRates.IMD") 
	'DeleteFile("Wires From Multiple Senders Over Threshold.IMD") 
	DeleteFile("Wires from Multiple Senders Over Thrsh - Details.IMD") 
	DeleteFile("Wires from Multiple Senders Over Thrsh - Details_tmp.IMD") 
	DeleteFile("Inward Wire Cust Acct Branch_US.IMD")
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

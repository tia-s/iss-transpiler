'====================================================================================================
'	Test#: 		WT15 - High Volume Drafts By Same Customer
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	18-Feb-14 12:30:04 AM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "WT15 - High Volume Drafts By Same Customer"
Const scriptname_log ="WT15 - High Volume Drafts By Same Customer.iss"
Global errors_string As String
Const division = "NCB"
Dim fso As Object


Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine
	
	client.closeall
	CreateComp_Month
	OutwardDrftsAccMastJoin
	SummarizeOutwardDrfts
	DetectMultipleMthlyDrfts
	JoinDetailsToSummary
	RenameFieldSummary
	Get_Final_Detail_File
	RenameFieldDetails
	Join_RiskRating
	ExportDatabaseSummary
	ExportDatabaseDetails
	Client.CloseAll
	
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

'------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------ANALYSIS-----------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------
' Append month field. Source file may cross overs two months.
Function CreateComp_Month
If haveRecords("Outward Draft FX.IMD") Then
	Set db = Client.OpenDatabase("Outward Draft FX.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "RUN_DATE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@Dtoc(@Date(),""YYYY-MM-DD"")"
		field.Length = 10
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
		
	Set field = db.TableDef.NewField
		field.Name = "MONTH"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@CompIf(@Month(BILL_DATE_DATE) = 1, ""JANUARY"", @Month(BILL_DATE_DATE) = 2, ""FEBRUARY"", @Month(BILL_DATE_DATE) = 3, ""MARCH"", @Month(BILL_DATE_DATE) = 4, ""APRIL"", @Month(BILL_DATE_DATE) = 5, ""MAY"", @Month(BILL_DATE_DATE) = 6, ""JUNE"", @Month(BILL_DATE_DATE) = 7, ""JULY"", @Month(BILL_DATE_DATE) = 8, ""AUGUST"", @Month(BILL_DATE_DATE) = 9, ""SEPTEMBER"", @Month(BILL_DATE_DATE) = 10, ""OCTOBER"", @Month(BILL_DATE_DATE) = 11, ""NOVEMBER"", @Month(BILL_DATE_DATE) = 12, ""DECEMBER"", 1,"""")"
		field.Length = 20
	If FieldExist(db.name, field.name) Then Call Delete_Virtual_Field(db.name, field.name)
		task.AppendField field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End If
End Function

' File: Join to get customer master information
Function OutwardDrftsAccMastJoin
If haveRecords("Outward Draft FX.IMD") And haveRecords("General Acct Master Lite.IMD")  Then
	Set db = Client.OpenDatabase("Outward Draft FX.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "General Acct Master Lite.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "FORACID"
		task.AddSFieldToInc "ACCT_NAME"
		task.AddSFieldToInc "CUST_ID"
		task.AddSFieldToInc "CUSTOMER_TYPE"
		task.AddSFieldToInc "CUST_NAME"
		task.AddSFieldToInc "PAN_GIR_NUM"
		task.AddMatchKey "OPER_ACID", "ACID", "A"
		task.CreateVirtualDatabase = False
		dbName = "Outward Drft Mthly Gen Acct.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Count the number of drafts sent
Function SummarizeOutwardDrfts
If haveRecords("Outward Drft Mthly Gen Acct.IMD") Then
	Set db = Client.OpenDatabase("Outward Drft Mthly Gen Acct.IMD")
	Set task = db.Summarization
		task.AddFieldToSummarize "CUST_ID"
		task.AddFieldToSummarize "MONTH"
		task.AddFieldToTotal "BILL_AMT_INR"
		task.AddFieldToInc "CUST_NAME"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "RUN_DATE"
		task.AddFieldToInc "PAN_GIR_NUM"
		task.Criteria = "CUST_ID <> """""
		dbName = "Summ Outward Drft Gen Acct Cust.IMD"
		task.OutputDBName = dbName
		task.CreatePercentField = FALSE
		task.StatisticsToInclude = SM_SUM
		task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Extract customers with multiple out drafts
Function DetectMultipleMthlyDrfts
If haveRecords("Summ Outward Drft Gen Acct Cust.IMD") Then
	Set db = Client.OpenDatabase("Summ Outward Drft Gen Acct Cust.IMD")
	Set task = db.Extraction
		task.IncludeAllFields
		dbName = "High Volume Outward Drafts_tmp.IMD"
		'task.AddExtraction dbName, "", " NO_OF_RECS  >=" & e_Wires_Out_Count
		task.AddExtraction dbName, "", ""
		task.CreateVirtualDatabase = False
		task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

' File: Collect detail records to be exported
Function JoinDetailsToSummary
If haveRecords("Outward Drft Mthly Gen Acct.IMD") And haveRecords("High Volume Outward Drafts.IMD")Then
	Set db = Client.OpenDatabase("Outward Drft Mthly Gen Acct.IMD")
	Set task = db.JoinDatabase
		task.FileToJoin "High Volume Outward Drafts_tmp.IMD"
		task.IncludeAllPFields
		task.AddSFieldToInc "CUST_ID"
		task.AddMatchKey "CUST_ID", "CUST_ID", "A"
		task.AddMatchKey "MONTH", "MONTH", "A"
		task.CreateVirtualDatabase = False
		dbName = "High Volume Outdraft Detail2.IMD"
		task.PerformTask dbName, "", WI_JOIN_MATCH_ONLY
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("TBAADM.FEX_CLEAN_INST_TABLE.IMD") And haveRecords("High Volume Outdraft Detail2.IMD") Then	
	Set db = Client.OpenDatabase("TBAADM.FEX_CLEAN_INST_TABLE.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "High Volume Outdraft Detail2.IMD"
			task.AddPFieldToInc "INSTRMNT_NUM"
			task.AddPFieldToInc "INSTRMNT_TYPE"
			task.AddPFieldToInc "INSTRMNT_AMT"
			task.AddPFieldToInc "INSTRMNT_DATE_DATE"
			task.IncludeAllSFields
			task.AddMatchKey "BILL_ID", "BILL_ID", "A"
			task.CreateVirtualDatabase = False
			task.Criteria = "DEL_FLG == ""N"""
			dbName = "High Volume Outdraft Detail Temp1.IMD"
			task.PerformTask dbName, "",  WI_JOIN_MATCH_ONLY
		Set task = Nothing
	Set db = Nothing
End If	
If haveRecords("High Volume Outdraft Detail2.IMD") Then
		Set db = Client.OpenDatabase("High Volume Outdraft Detail2.IMD")
		Set task = db.JoinDatabase
			task.FileToJoin "TBAADM.FEX_CLEAN_INST_TABLE.IMD"
			task.AddSFieldToInc "INSTRMNT_NUM"
			task.AddSFieldToInc "INSTRMNT_TYPE"
			task.AddSFieldToInc "INSTRMNT_AMT"
			task.AddSFieldToInc "INSTRMNT_DATE_DATE"
			task.IncludeAllPFields
			task.AddMatchKey "BILL_ID", "BILL_ID", "A"
			task.CreateVirtualDatabase = False
			task.Criteria = ""
			dbName = "High Volume Outdraft Detail Temp2.IMD"
			task.PerformTask dbName, "",  WI_JOIN_NOC_SEC_MATCH
		Set task = Nothing
	Set db = Nothing
End If
End Function

Function Get_Final_Detail_File
  Client.CloseAll

  Dim FileOneCount As Long
  Dim FileTwoCount As Long

' Count records in both files
  If haveRecords("High Volume Outdraft Detail Temp1.IMD") Then
  Set db1 = Client.OpenDatabase("High Volume Outdraft Detail Temp1.IMD")	'---File One
         FileOneCount = db1.count
       Else
         FileOneCount = 0
  End If
  
  If haveRecords("High Volume Outdraft Detail Temp2.IMD") Then
  Set db2 = Client.OpenDatabase("High Volume Outdraft Detail Temp2.IMD")	'---File Two
     FileTwoCount = db2.count
     Else
              FileTwoCount = 0
  End If
     
  Set db1 = Nothing
  Set db2 = Nothing

' Test files for records before appending


      If FileOneCount = 0 And FileTwoCount  = 0 Then
        	Client.CloseAll
	Set task = Nothing
	Set db = Nothing


      ElseIf FileOneCount > 0 And FileTwoCount  = 0 Then
        	Client.CloseAll
        	Set ProjectManagement = client.ProjectManagement
	  ProjectManagement.RenameDatabase "High Volume Outdraft Detail Temp1.IMD", "High Volume Outdraft Detail"
	Set ProjectManagement = Nothing
		
       ElseIf FileOneCount = 0 And FileTwoCount  > 0 Then
         	Client.CloseAll
  	Set ProjectManagement = client.ProjectManagement
	  ProjectManagement.RenameDatabase "High Volume Outdraft Detail Temp2.IMD", "High Volume Outdraft Detail"
	Set ProjectManagement = Nothing
	
       Else ' Append...No file is empty
       	Set db = Client.OpenDatabase("High Volume Outdraft Detail Temp2.IMD")
	Set task = db.AppendDatabase
		task.AddDatabase "High Volume Outdraft Detail Temp1.IMD"
		dbName = "High Volume Outdraft Detail.IMD"
		task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
    End If

  Set db1 = Nothing
  Set db2 = Nothing
End Function


Function RenameFieldSummary
If haveRecords("High Volume Outward Drafts_tmp.IMD") Then
	Set db = Client.OpenDatabase("High Volume Outward Drafts_tmp.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_ID"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 9
		task.ReplaceField "CUST_ID", field

	Set field = db.TableDef.NewField
		field.Name = "NO_OF_WIRES"
		field.Description = "Number of records found for this key value"
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 0
		task.ReplaceField "NO_OF_RECS", field

	Set field = db.TableDef.NewField
		field.Name = "DRAFTS_TOTAL_JMD"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT_INR_SUM", field

	Set field = db.TableDef.NewField
		field.Name = "CUSTOMER_NAME"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 80
		task.ReplaceField "CUST_NAME", field

	Set field = db.TableDef.NewField
		field.Name = "TRN"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 25
		task.ReplaceField "PAN_GIR_NUM", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function RenameFieldDetails
If haveRecords("High Volume Outdraft Detail.IMD") Then
	Set db = Client.OpenDatabase("High Volume Outdraft Detail.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "DRAFT_AMOUNT_JMD"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT_INR", field
		

	Set field = db.TableDef.NewField
		field.Name = "DRAFT_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "DATE_OF_REMIT_DATE", field

	Set field = db.TableDef.NewField
		field.Name = "CURRENCY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 3
		task.ReplaceField "BILL_CRNCY_CODE", field

	Set field = db.TableDef.NewField
		field.Name = "DRAFT_AMOUNT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 4
		task.ReplaceField "BILL_AMT", field

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
		field.Name = "INSTRUMENT_NUM"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 16
		task.ReplaceField "INSTRMNT_NUM", field

	Set field = db.TableDef.NewField
		field.Name = "INSTRUMENT_TYPE"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""
		field.Length = 5
		task.ReplaceField "INSTRMNT_TYPE", field

	Set field = db.TableDef.NewField
		field.Name = "INSTRUMENT_AMT"
		field.Description = ""
		field.Type = WI_NUM_FIELD
		field.Equation = ""
		field.Decimals = 2
		task.ReplaceField "INSTRMNT_AMT", field

	Set field = db.TableDef.NewField
		field.Name = "INSTRUMENT_DATE"
		field.Description = ""
		field.Type = WI_DATE_FIELD
		field.Equation = "YYYYMMDD"
		task.ReplaceField "INSTRMNT_DATE_DATE", field
		task.PerformTask
	Set field = Nothing
	Set task = Nothing
	Set db = Nothing
End if
End Function

Function Join_RiskRating
If haveRecords("Customer_Turnover_wRisk.IMD") And haveRecords("High Volume Outward Drafts_tmp.IMD") Then
	Set db = Client.OpenDatabase("High Volume Outward Drafts_tmp.IMD")	
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
		dbName = "High Volume Outward Drafts.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function ExportDatabaseSummary
If haveRecords("High Volume Outward Drafts.IMD") Then
	Set db = Client.OpenDatabase("High Volume Outward Drafts.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "MONTH"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "NO_OF_WIRES"
		task.AddFieldToInc "DRAFTS_TOTAL_JMD"
		task.AddFieldToInc "TRN"
		task.AddFieldToInc "RUN_DATE"
		task.AddFieldToInc "MONTHLY_DEPOSIT"
		task.AddFieldToInc "RISK_SCORE"
		task.AddFieldToInc "RISK_CATEGORY"
		task.AddFieldToInc "OCCUPATION"
		task.AddFieldToInc "INDUSTRY"
		task.AddFieldToInc "ANNUAL_INCOME"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\WT15_SUMM.MDB", "Database", "MDB2000", 1, db.Count, eqn
		RESULTSLOG(db.name)
	Set task = Nothing
	Set db = Nothing

Else 
NORESULTSLOG("WT15_SUMMARY.IMD") 
End If
	
End Function

Function ExportDatabaseDetails
If haveRecords("High Volume Outdraft Detail.IMD") Then
	Set db = Client.OpenDatabase("High Volume Outdraft Detail.IMD")
	Set task = db.ExportDatabase
		task.AddFieldToInc "CUSTOMER_ID"
		task.AddFieldToInc "CUSTOMER_NAME"
		task.AddFieldToInc "ACCOUNT_NUMBER"
		task.AddFieldToInc "ACCOUNT_NAME"
		task.AddFieldToInc "MONTH"
		task.AddFieldToInc "BILL_ID"
		task.AddFieldToInc "INSTRUMENT_NUM"
		task.AddFieldToInc "INSTRUMENT_TYPE"
		task.AddFieldToInc "INSTRUMENT_AMT"
		task.AddFieldToInc "INSTRUMENT_DATE"
		task.AddFieldToInc "DRAFT_DATE"
		task.AddFieldToInc "CURRENCY"
		task.AddFieldToInc "DRAFT_AMOUNT"
		task.AddFieldToInc "DRAFT_AMOUNT_JMD"
		task.AddFieldToInc "OTHER_PARTY_NAME"
		task.AddFieldToInc "CUSTOMER_TYPE"
		task.AddFieldToInc "RUN_DATE"
		eqn = ""
		task.PerformTask Client.WorkingDirectory & "Reports\WT15_DTLS.MDB", "Database", "MDB2000", 1, db.Count, eqn
	Set task = Nothing
	Set db = Nothing
End If
End Function



Function CleanUp
  'DeleteFile("High Volume Outward Drafts.IMD")
  DeleteFile("High Volume Outward Drafts_tmp.IMD")
  DeleteFile("Summ Outward Drft Gen Acct Cust.IMD")
  DeleteFile("Outward Drft Mthly Gen Acct.IMD")
  DeleteFile("High Volume Outdraft Detail.IMD")
  DeleteFile("High Volume Outdraft Detail2.IMD")
  DeleteFile ("High Volume Outdraft Detail Temp2.IMD")
  DeleteFile ("High Volume Outdraft Detail Temp1.IMD")
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

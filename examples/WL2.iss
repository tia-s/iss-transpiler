'====================================================================================================
'	Test# 		WL2
'	Risk		
' 	Objective	
' 	Frequency	
' 	Last Modified	12282014 35821 PM
'====================================================================================================
'	Script Dependencies
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = WL2
Const scriptname_log =WL2.iss
Global errors_string As String
Const division = UTC
Dim fso As Object

Sub Main
	Ignorewarning(True)
	Set fso    = CreateObject(Scripting.FileSystemObject)
	On Error GoTo finalRoutine

Client.CloseAll

	Call ABJoinTransWatchlist()
	Call ACRenameFields()
	Call BSummDetails()		
	Call FRename_ResultFields()
	Call FBExtResult_Summ()
	Call FCJoinExtResult_Dtl()
	Call GExportDatabase()
'	Call JCleanUp()
	
	finalRoutine
	
	If err.description   Or err.number  0 Then
		message  = err.description &   & err.number
		info = Error
	Else
		message = Script Completed Successfully
		info = Information
	End If	
	
	Call logfile(scriptname_log, End, Data Analysis, info, message & errors_string)


	Client.CloseAll
	Client.Quit
End Sub


'---------------------------------------------------------------------------------------------------------------------------------------
' Logfile(ByVal Log_Step As String, ByVal Log_Step As String, ByVal Log_Action As String, ByVal Log_MsgType As String, ByVal Log_Message As String)
'
' Input
' filename_log		- {String} Which file the analysis is on
' Log_Step		- {String} Which step is being run
' Log Action		- {String} Which Action is performed
' Log_Msgtype		- {String} Log Type (Informational, Error, Warning)
' Log_Message		- {String} Log Message
'
' Returns 		Nothing
'
' Description This function creates and appends to a logfile 
'---------------------------------------------------------------------------------------------------------------------------------------	
Function logfile(ByVal filename_log As String,  ByVal Log_Step As String, ByVal Log_Action As String, ByVal Log_MsgType As String, ByVal Log_Message As String)

On Error GoTo exit_logfile

If e_debug  True Then Exit Sub

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


If (Len(e_logfilename)  0) Then logfilename = e_logfilename & .imd Else logfilename = log_file.imd

	'Create the table if it doesn't exist. 
	Set pm = Client.ProjectManagement
	If Not pm.DoesDatabaseExist(logfilename) Then
		Set NewTable = Client.NewTableDef
		Set AddedField = NewTable.NewField
		AddedField.Name = LOG_DATE
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = LOG_TIME
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False				
		AddedField.Name = FILENAME
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = SCRIPTNAME
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=50	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = SCRIPTSTEP
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=50	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = ACTION
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		AddedField.Name = MSG_TYPE
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False		
		AddedField.Name = MESSAGE
		AddedField.Type = WI_CHAR_FIELD
		AddedField.length=500	
		NewTable.AppendField AddedField
		NewTable.Protect = False
		Set db1 = Client.NewDatabase(Logfilename, , NewTable)
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

		rec1.setcharvalueat 1, Format (Now(), Short Date)
		rec1.setcharvalueat 2, Format (Now(), Medium Time)
		If filename_log   Then 	rec1.setcharvalueat 3, filename_log
		If scriptname_log   Then rec1.setcharvalueat 4, scriptname_log
		If Log_Step    Then rec1.setcharvalueat 5, Log_Step   
		If Log_Action   Then rec1.setcharvalueat 6, Log_Action
		If Log_MsgType   Then rec1.setcharvalueat 7, Log_MsgType
		If Log_Message   Then rec1.setcharvalueat 8, Log_Message
		
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
	
exit_logfile	
	
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
				If rs.count  0 Then
					haveRecords = True
				Else
					errors_string = errors_string &  with errors - & dbname &  has no records. & Chr(10)
					Call logfile(dbname, division, haverecords, Error, Database does not have records.)
				End If
			Set rs = Nothing
		db.close
		Set db = Nothing
	Else
		errors_string = errors_string &  with errors - & dbname &  missing. & Chr(10)
		Call logfile(dbname, division, haverecords, Error, Database does not exist.)
		
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

' Join Today's Transactions to Internal Watch List

Function ABJoinTransWatchlist
If haveRecords(Daily_Transactions_Today.IMD) And haveRecords(AML Compliance Internal Monitoring List- & Trim(e_Internal_List_Sheet_Name) & .IMD)  Then
	Set db = Client.OpenDatabase(Daily_Transactions_Today.IMD)
	Set task = db.JoinDatabase
	task.FileToJoin AML Compliance Internal Monitoring List- & Trim(e_Internal_List_Sheet_Name) & .IMD  
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey UTCID, UTC_ID, A
	task.Criteria =  UTCID   
	task.CreateVirtualDatabase = False
	dbName = WL2_DTLS_INT.IMD
	task.PerformTask dbName, , WI_JOIN_MATCH_ONLY
	
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function ACRenameFields
If haveRecords(WL2_DTLS_INT.IMD) Then
Set db = Client.OpenDatabase(WL2_DTLS_INT.IMD)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = WL_RISK_RATING
	field.Description = 
	field.Type = WI_CHAR_FIELD
	field.Equation = 
	field.Length = 15
	task.ReplaceField RISK_RATING1, field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = WL_RISK_COMMENT
	field.Description = 
	field.Type = WI_CHAR_FIELD
	field.Equation = 
	field.Length = 81
	task.ReplaceField RISK_COMMENT, field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function


' Summarize WL2 Details
Function BSummDetails
If haveRecords(WL2_DTLS_INT.IMD) Then
Set db = Client.OpenDatabase(WL2_DTLS_INT.IMD)
	Set task = db.Summarization
	task.AddFieldToSummarize UTCID
	task.AddFieldToSummarize POST_DATE
	task.AddFieldToInc CUSTOMER_BRANCH
	task.AddFieldToInc HOLDER_TYPE
	task.AddFieldToInc CATEGORY
	task.AddFieldToInc WL_RISK_RATING
	task.AddFieldToInc WL_RISK_COMMENT
	task.AddFieldToInc NAME_OF_CUSTOMER
	task.AddFieldToInc OCCUPATION
	task.AddFieldToInc RATING_SOURCE
	task.AddFieldToInc RISK_RATING
	task.AddFieldToInc BRANCH_NAME	
	task.AddFieldToTotal TRANSACTION_AMOUNT
	dbName = WL2_SUMM_INT.IMD
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function FRename_ResultFields
If haveRecords(WL2_SUMM_INT.IMD) Then
Set db = Client.OpenDatabase(WL2_SUMM_INT.IMD)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = NO_OF_TRANSACTIONS
	field.Description = Number of records found for this key value
	field.Type = WI_NUM_FIELD
	field.Equation = 
	field.Decimals = 0
	task.ReplaceField NO_OF_RECS, field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = TRANSACTION_SUMMARY
	field.Description = 
	field.Type = WI_NUM_FIELD
	field.Equation = 
	field.Decimals = 2
	task.ReplaceField TRANSACTION_AMOUNT_SUM, field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
End Function

'April 20 - Include new thresholds for Medium and High
	
Function FBExtResult_Summ
If haveRecords(WL2_Summ_INT.IMD) Then
	Set db = Client.OpenDatabase(WL2_Summ_INT.IMD)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = WL2_SUMM.IMD
	task.AddExtraction dbName, , (WL_RISK_RATING == High .AND. TRANSACTION_SUMMARY =  & e_WL2_High_Thresh & ) .OR. WL_RISK_RATING == Medium .AND. TRANSACTION_SUMMARY =  & e_WL2_Medium_Thresh 
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function FCJoinExtResult_Dtl
If haveRecords(Daily_Transactions_Today.IMD) Then
	Set db = Client.OpenDatabase(Daily_Transactions_Today.IMD)
	Set task = db.JoinDatabase
   If haveRecords(WL2_Summ.IMD) Then
	task.FileToJoin WL2_Summ.IMD
	task.IncludeAllPFields
	task.AddSFieldToInc NO_OF_TRANSACTIONS
	task.AddMatchKey UTCID, UTCID, A
	task.CreateVirtualDatabase = False
	dbName = WL2_DTLS.IMD
	task.PerformTask dbName, , WI_JOIN_MATCH_ONLY
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function GExportDatabase
If haveRecords(WL2_DTLS.IMD) Then
Set db = Client.OpenDatabase(WL2_DTLS.IMD)
	Set task = db.ExportDatabase
	task.AddFieldToInc UTCID
	task.AddFieldToInc ACCT_NO
	task.AddFieldToInc TRUST_CODE
	task.AddFieldToInc RECEIPT_NUMBER
	task.AddFieldToInc TRANSACTION_BRANCH
	task.AddFieldToInc BRANCH_DESCRIPTION
	task.AddFieldToInc TRANSACTION_TYPE
	task.AddFieldToInc TRAN_DATE
	task.AddFieldToInc TRANSACTION_AMOUNT
	task.AddFieldToInc UNITS
	task.AddFieldToInc UNIT_PRICE
	task.AddFieldToInc PAYMENT_TYPE
	task.AddFieldToInc TRANSACTION_CHANNEL
	task.AddFieldToInc POST_DATE
	task.AddFieldToInc AGENT_CODE
	task.AddFieldToInc CUSTOMER_NAME
	task.AddFieldToInc ADDRESS1
	task.AddFieldToInc ADDRESS2
	task.AddFieldToInc ADDRESS3
	task.AddFieldToInc DATE_OF_BIRTH
	task.AddFieldToInc OCCUPATION
	task.AddFieldToInc RATING_SOURCE
	task.AddFieldToInc RISK_RATING
	task.AddFieldToInc BRANCH_NAME
	task.AddFieldToInc HOLDER_NAME
	task.AddFieldToInc NARRATIVE
	task.AddFieldToInc HOLDER_TYPE
	eqn = 
	task.PerformTask Client.WorkingDirectory & ReportsWL2_DTLS.MDB, Database, MDB2000, 1, db.Count, eqn
	Set db = Nothing
End If

If haveRecords(WL2_SUMM.IMD) Then
	Set db = Client.OpenDatabase(WL2_SUMM.IMD)
	Set task = db.ExportDatabase
	task.AddFieldToInc UTCID
	task.AddFieldToInc CUSTOMER_BRANCH
	task.AddFieldToInc HOLDER_TYPE
	task.AddFieldToInc POST_DATE
	task.AddFieldToInc NO_OF_TRANSACTIONS
	task.AddFieldToInc TRANSACTION_SUMMARY
	task.AddFieldToInc CATEGORY
	task.AddFieldToInc WL_RISK_RATING
	task.AddFieldToInc WL_RISK_COMMENT
	task.AddFieldToInc NAME_OF_CUSTOMER
	task.AddFieldToInc OCCUPATION
	task.AddFieldToInc RATING_SOURCE
	task.AddFieldToInc RISK_RATING
	task.AddFieldToInc BRANCH_NAME
	eqn = 
	task.PerformTask Client.WorkingDirectory & ReportsWL2_SUMM.MDB, Database, MDB2000, 1, db.Count, eqn
	Set db = Nothing
End If
	Set task = Nothing
End Function


Function JCleanUp
	DeleteFile(WL2_SUMM.IMD)
	DeleteFile(WL2_DTLS.IMD)
	DeleteFile(WL2_SUMM_INT.IMD)
	DeleteFile(WL2_DTLS_INT.IMD)
End Function


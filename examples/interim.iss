'----- Constants -----

Dim v_CurrentDate As Date
Dim v_CurrentDateChar As String

Dim v_AverageDateStart As Date
Dim v_AverageDateStartChar As String

Dim v_AverageDateEnd As Date
Dim v_AverageDateEndChar As String

'====================================================================================================
'	Test#: 		Interim
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	12/28/2014 3:55:55 PM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "Interim"
Const scriptname_log ="Interim.iss"
Global errors_string As String
Const division = "UAT"
Dim fso As Object

Sub Main

Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
	On Error GoTo finalRoutine
Client.CloseAll

' Use date fields to get current date, inactive history and average history start date.
	v_CurrentDate = CDate(CLng(Date) - 3)
	v_AverageDateStart = CDate(CLng(Date) - e_TM1_AVG_Days - 3)
	v_AverageDateEnd = CDate(CLng(Date) - 3)
	
'	v_CurrentDate = "2022-04-08"
'	v_AverageDateStart = "2021-04-08"
'	v_AverageDateEnd = "2022-04-08"
	
	v_CurrentDateChar = Format(v_CurrentDate, "YYYYMMDD")
	v_AverageDateStartChar = Format(v_AverageDate, "YYYYMMDD")
	v_AverageDateEndChar = Format(v_AverageDateEnd, "YYYYMMDD")

 
	
'Run Scripts	
	Call AddFields()
Client.CloseAll
	Call AExt_Master()
 
	Call BACreateDate()
	Call BAJoinCustomers()
 Client.CloseAll
	Call BACreateKYCFields()
	Call BBCreateKYCFields()
Client.CloseAll

	Call CCreate_Fields()
	 
	Call DRename_Fields()
 Client.CloseAll
	Call EExt_History()
 
	Call FAppend_History()
 Client.CloseAll
 
	Call GJoin_Hist_Master()
 
	Call HJoinCustomers()	
 Client.CloseAll
	Call IJoinBranch()
	
	Call JCreate_Fields()
	Call KRename_SourceFields()
 Client.CloseAll
	Call LCreate_DateFields()
	
 Client.CloseAll
	Call MExt_Physical()
 Client.CloseAll
  
	Call MFJoinPortaltoRisk()
Client.CloseAll	 
	Call MG_CreateKeys()
	Call MHJoin_Portal()
Client.CloseAll	

	Call MI_JoinCustomerBranch()
 
 Client.CloseAll
	Call NExtHist_Daily() 		
	Call OExtHist_Average()	
 Client.CloseAll	
	Call QCreateAccountDate()
	Call RExtAccounts_Daily()
'	Call SCleanup()
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
	Client.Quit
End Sub


' Append Field
Function AddFields

If haveRecords("REPORTSVR.VDPUTR04N.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR04N.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 15
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 	

If haveRecords("REPORTSVR.VDPUTR081.IMD") Then	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR081.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 12
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 

If haveRecords("REPORTSVR.VDPUTR021IMD") Then	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR021.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 15
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If

If haveRecords("REPORTSVR.VDPUTR011.IMD") Then	
	
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTR011.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 12
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If

If haveRecords("REPORTSVR.VDPUCID01.IMD") Then	
	
	
		Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BIRTHDAT"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "BIRTHDAT_DATE", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End If 
End Function

Function AExt_Master
If haveRecords("REPORTSVR.VDPUTP01.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTP01.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "ACCTNOP"
	dbName = "Master_Funds.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function BACreateDate
If haveRecords("REPORTSVR.VDPUCID01.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE_CREATED"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(@Left(@AllTrim(RECDATE1_DATE),8),""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask
End If
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function


Function BAJoinCustomers
If haveRecords("dbo.Registrations.IMD") Then
Set db = Client.OpenDatabase("dbo.Registrations.IMD")
	Set task = db.JoinDatabase
   If haveRecords("REPORTSVR.VDPUCID01.IMD") Then
	task.FileToJoin "REPORTSVR.VDPUCID01.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ADDRESS1"
	task.AddSFieldToInc "ADDRESS2"
	task.AddSFieldToInc "ADDRESS3"
	task.AddSFieldToInc "BIRTHDAT"
	task.AddSFieldToInc "HOLDRTYP"
	task.AddSFieldToInc "FIRSTNAME"
	task.AddSFieldToInc "SURNAME"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Registrations.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function


Function BACreateKYCFields
If haveRecords("Registrations.IMD") Then
	Set db = Client.OpenDatabase("Registrations.IMD")
		Set task  = db.TableManagement
'Field created on March 21, 2015 to remove data errors. 		
	Set field = db.TableDef.NewField
		field.Name = "PLACEOFWORK_NEW"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@IF(@List(@UPPER(PLACEOFWORK), ""N.A."", ""N/A"", ""N\A"", ""Private"", ""Not APPLICABLE"", ""NA"" , ""NONE"", ""UNEMPLOYED"", ""UNKNOWN""), "" "", PLACEOFWORK)"
		field.Length = 50
		task.AppendField field
		task.PerformTask
'Field created on March 21, 2015 to remove data errors. 	
	Set field = db.TableDef.NewField
		field.Name = "EMPLOYERNAME_NEW"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@IF(@List(@UPPER(EMPLOYERNAME), ""N.A."", ""N/A"", ""N\A"", ""Private"", ""Not APPLICABLE"", ""NA"" , ""NONE"", ""UNEMPLOYED"", ""UNKNOWN""), "" "", EMPLOYERNAME)"
		field.Length = 50
		task.AppendField field
		task.PerformTask
End If
	Set task = Nothing
	Set db = Nothing
End Function


'Computes which KYC information is missing from records
'Commented equations reflect changes based on onsite UAT review in April 2015.
Function BBCreateKYCFields
If haveRecords("Registrations.IMD") Then
	Set db = Client.OpenDatabase("Registrations.IMD")
		Set task  = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "DATE_OF_BIRTH"
		field.Description = ""
		field.Type = WI_VIRT_DATE
		field.Equation = "BIRTHDAT"
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_FNAME"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim(FIRSTNAME) == """" .AND. HOLDRTYP == ""I"", ""First Name"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask	
	Set field = db.TableDef.NewField
		field.Name = "CHECK_SNAME"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@CompIf(@Alltrim(SURNAME) == """" .AND. HOLDRTYP == ""I"", ""Last Name"", @Alltrim(SURNAME) == """",""Business Name"",1,"""")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_DOB"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@Compif(@ALLTRIM(@STR(BIRTHDAT, 7,0)) == ""0"" .AND. HOLDRTYP <> ""C"", ""Date of Birth"", @Alltrim(@Str(BIRTHDAT, 7,0)) == ""0"" ,""Incorporation Date"", 1, """")"
		field.Equation = "@Compif(DATE_OF_BIRTH > @Date() .AND. HOLDRTYP <> ""C"", ""Date of Birth"",  DATE_OF_BIRTH > @Date() ,""Incorporation Date"", 1, """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_SEX"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim(SEX) == """", ""Gender"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_ADDRESS1"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( ADDRESS1) == """", ""Street Address"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_ADDRESS2"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( ADDRESS2) == """", ""Address City"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
' Field updated March 21,2015 to exclude Check Address 3 if Check_COR is not blank.		
		Set field = db.TableDef.NewField
		field.Name = "CHECK_ADDRESS3"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@if(@Alltrim(ADDRESS3) == """", ""Address Country"", """")"
		field.Equation = "@if(@Alltrim(COUNTRYOFRESIDENCE) == """"  .AND. ADDRESS3 == """", ""Address Country"" , """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_ID"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@CompIf(@Alltrim(IDNO) == """" .AND. HOLDRTYP == ""I"", ""ID"", @Alltrim(IDNO) == """",""Company ID"",1,"""")"
		field.Equation = "@CompIf(@Alltrim(IDNO) == """"  .AND. @AllTrim(PPNO) =="""" .AND. @AllTrim(DPNO)== """" .AND.  HOLDRTYP == ""I"", ""ID"", @Alltrim(IDNO) == """" .AND. @AllTrim(PPNO) =="""" .AND. @AllTrim(DPNO)== """",""ID"",1,"""")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_IDEXPIRE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@if(@Alltrim( IDNOEXPIRE_DATE) == ""00000000"" .AND. HOLDRTYP == ""I"", ""ID Expiry Date"", """")"
		field.Equation = "@if(IDNOEXPIRE_DATE < @Date() .AND. DPNOEXPIRE_DATE < @Date() .AND. PPNOEXPIRE_DATE < @Date() .AND. HOLDRTYP == ""I"", ""ID Expiry Date"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_HOMEOWNERSHIP"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( HOMEOWNERSHIP) == """" .AND. HOLDRTYP == ""I"", ""Home Ownership"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_INCOME"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
	'	field.Equation = "@CompIf(@Alltrim(ANNUALINCOME) == """" .AND. HOLDRTYP == ""I"", ""Annual Income"", @Alltrim( ANNUALINCOME) == """",""Annual Income"",1,"""")"
		field.Equation = "@CompIf(ANNUALINCOME ==0  .AND. HOLDRTYP == ""I"", ""Annual Income"", ANNUALINCOME == 0,""Annual Income"",1,"""")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_MULTIPLECITIZENSHIP"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( COUNTRYMULTIPLECITIZENSHIP) == """" .AND. HOLDRTYP == ""I"", ""Multiple Citizenship"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_COB"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( COUNTRYOFBIRTH) == """", ""Country of Birth"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_CITIZENSHIP"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@CompIf(@Alltrim(COUNTRYOFCITIZENSHIP) == """" .AND. HOLDRTYP == ""I"", ""Country of Citizenship"", @Alltrim(COUNTRYOFCITIZENSHIP) == """",""Country of Citizenship"",1,"""")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_EMPLOYMENT"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( EMPLOYMENTSTATUS) == """", ""Employment Status"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_OCCUPATION"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim( OCCUPATION) == """", ""Occupation"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_PHONE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@if(@Alltrim( PHONEHOME) == """", ""Phone Number"", """")"
		field.Equation = "@if(@Alltrim( PHONEHOME) == """" .AND. @AllTrim( PHONEMOBILE) =="""" .AND. @AllTrim( WORKCONTACTNO) =="""", ""Phone Number"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_RELATIONSHIP"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim(PURPOSEOFRELATIONSHIP) == """", ""Purpose of Relationship"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	 Set field = db.TableDef.NewField
		field.Name = "CHECK_POC"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@if(@Alltrim(PLACEOFWORK_NEW) == """", ""Place of Work"", """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.Name = "CHECK_EMPLOYER"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@if(@Alltrim( EMPLOYERNAME) == """", ""Employer Name"", """")"
		field.Equation = "@if(@Alltrim(PLACEOFWORK_NEW) == """" .AND. EMPLOYMENTSTATUS == ""Employed"" .AND. @AllTrim(EMPLOYERNAME_NEW) == """", ""Employer Name"" , """")"
		field.Length = 50
		task.AppendField field
		task.PerformTask
	Set field = db.TableDef.NewField
		field.name = "MISSING_KYC_IND"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
'		field.Equation = "@AllTrim(CHECK_FNAME) + "" "" + @AllTrim(CHECK_SNAME) + "" ""  + @AllTrim(CHECK_DOB) + "" ""  + @AllTrim(CHECK_SEX) + "" ""  + @AllTrim(CHECK_ADDRESS1)+ "" ""  + @AllTrim(CHECK_ADDRESS2) + "" "" + @AllTrim(CHECK_ADDRESS3) + "" ""  + @AllTrim(CHECK_ID) + "" ""  + @AllTrim(CHECK_IDEXPIRE) + "" ""  +@AllTrim(CHECK_HOMEOWNERSHIP) + "" ""  + @AllTrim(CHECK_INCOME) + "" ""  + @AllTrim(CHECK_MULTIPLECITIZENSHIP) + "" ""  + @AllTrim(CHECK_COB) + "" ""  + @AllTrim(CHECK_COR) + "" ""  + @AllTrim(CHECK_EMPLOYMENT) + "" ""  + @AllTrim(CHECK_CITIZENSHIP) + "" ""  + @AllTrim(CHECK_OCCUPATION) + "" ""  + @AllTrim(CHECK_EMPLOYER) + "" "" + @AllTrim(CHECK_PHONE) + "" ""  + @AllTrim(CHECK_RELATIONSHIP) + "" ""  + @AllTrim(CHECK_POC )"
'		field.Equation = "@AllTrim(CHECK_FNAME) + "" "" + @AllTrim(CHECK_SNAME) + "" ""  + @AllTrim(CHECK_DOB) + "" ""  + @AllTrim(CHECK_SEX) + "" ""  + @AllTrim(CHECK_ADDRESS1)+ "" ""  + @AllTrim(CHECK_ADDRESS2) + "" "" + @AllTrim(CHECK_ADDRESS3) + "" ""  + @AllTrim(CHECK_ID) + "" ""  + @AllTrim(CHECK_IDEXPIRE) + "" ""  +@AllTrim(CHECK_HOMEOWNERSHIP) + "" ""  + @AllTrim(CHECK_INCOME) + "" ""  + @AllTrim(CHECK_MULTIPLECITIZENSHIP) + "" ""  + @AllTrim(CHECK_COB) + "" ""  + @AllTrim(CHECK_COR) + "" ""  + @AllTrim(CHECK_EMPLOYMENT) + "" ""  + @AllTrim(CHECK_CITIZENSHIP) + "" ""  + @AllTrim(CHECK_OCCUPATION) + "" ""  + @AllTrim(CHECK_EMPLOYER) + "" "" + @AllTrim(CHECK_PHONE) + "" ""  + @AllTrim(CHECK_RELATIONSHIP)"
'		field.Equation = "@AllTrim(CHECK_FNAME) + "" "" + @AllTrim(CHECK_SNAME) + "" ""  + @AllTrim(CHECK_DOB) + "" ""  + @AllTrim(CHECK_SEX) + "" ""  + @AllTrim(CHECK_ADDRESS1)+ "" ""  + @AllTrim(CHECK_ADDRESS2) + "" "" + @AllTrim(CHECK_ADDRESS3_NEW) + "" ""  + @AllTrim(CHECK_ID) + "" ""  + @AllTrim(CHECK_IDEXPIRE) + "" ""  +@AllTrim(CHECK_HOMEOWNERSHIP) + "" ""  + @AllTrim(CHECK_INCOME) + "" ""  + @AllTrim(CHECK_MULTIPLECITIZENSHIP) + "" ""  + @AllTrim(CHECK_COB) + "" ""  + @AllTrim(CHECK_COR) + "" ""  + @AllTrim(CHECK_EMPLOYMENT) + "" ""  + @AllTrim(CHECK_CITIZENSHIP) + "" ""  + @AllTrim(CHECK_OCCUPATION) + "" ""  + @AllTrim(CHECK_EMPLOYER) + "" "" + @AllTrim(CHECK_PHONE) + "" ""  + @AllTrim(CHECK_RELATIONSHIP)"
'		field.Equation = "@AllTrim(CHECK_FNAME) + "" "" + @AllTrim(CHECK_SNAME) + "" ""  + @AllTrim(CHECK_DOB) + "" ""  + @AllTrim(CHECK_SEX) + "" ""  + @AllTrim(CHECK_ADDRESS1)+ "" ""  + @AllTrim(CHECK_ADDRESS2) + "" "" + @AllTrim(CHECK_ADDRESS3_NEW) + "" ""  + @AllTrim(CHECK_ID) + "" ""  + @AllTrim(CHECK_IDEXPIRE) + "" ""  +@AllTrim(CHECK_HOMEOWNERSHIP) + "" ""  + @AllTrim(CHECK_INCOME) + "" ""  + @AllTrim(CHECK_MULTIPLECITIZENSHIP) + "" ""  + @AllTrim(CHECK_COB) + "" ""  + @AllTrim(CHECK_COR) + "" ""  + @AllTrim(CHECK_EMPLOYMENT) + "" ""  + @AllTrim(CHECK_CITIZENSHIP) + "" ""  + @AllTrim(CHECK_OCCUPATION) + "" ""  + @AllTrim(CHECK_EMPLOYER_NEW) + "" "" + @AllTrim(CHECK_PHONE) + "" ""  + @AllTrim(CHECK_RELATIONSHIP)"
		field.Equation = "@AllTrim(CHECK_FNAME) + "" "" + @AllTrim(CHECK_SNAME) + "" ""  + @AllTrim(CHECK_DOB) + "" ""  + @AllTrim(CHECK_SEX) + "" ""  + @AllTrim(CHECK_ADDRESS1)+ "" ""  + @AllTrim(CHECK_ADDRESS2) + "" "" + @AllTrim(CHECK_ADDRESS3) + "" ""  + @AllTrim(CHECK_ID) + "" ""  + @AllTrim(CHECK_IDEXPIRE) + "" ""  +@AllTrim(CHECK_HOMEOWNERSHIP) + "" ""  + @AllTrim(CHECK_INCOME) + "" ""  + @AllTrim(CHECK_MULTIPLECITIZENSHIP) + "" ""  + @AllTrim(CHECK_COB) + "" ""  + @AllTrim(CHECK_EMPLOYMENT) + "" ""  + @AllTrim(CHECK_CITIZENSHIP) + "" ""  + @AllTrim(CHECK_OCCUPATION) + "" ""  + @AllTrim(CHECK_EMPLOYER) + "" "" + @AllTrim(CHECK_PHONE) + "" ""  + @AllTrim(CHECK_RELATIONSHIP)"
		field.Length = 100
		task.AppendField field
		task.PerformTask
Set field = Nothing
Set task = Nothing
Set db = Nothing	
End If
End Function


'Create Fields
Function CCreate_Fields
If haveRecords("REPORTSVR.VDPUCID01.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUCID01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "HOLDRNAME"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@ALLTRIM(FIRSTNAME) + "" "" + @ALLTRIM(SURNAME)"
	field.Length = 50
	task.AppendField field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DOB"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "BIRTHDAT"
	task.AppendField field
	task.PerformTask	
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
		
If haveRecords("REPORTSVR.VDPUTR04N.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR04N.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VALUNITS_New"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "VALUNITS * " & e_US_Rate & ""
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing

If haveRecords("REPORTSVR.VDPURF031.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPURF031.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRUSTCOD"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """005"""
	field.Length = 3
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

'Rename fields for Append
Function DRename_Fields
If haveRecords("REPORTSVR.VDPURF031.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPURF031.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RECEIPTNO"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 16
	task.ReplaceField "TRNREF", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	
'	Set field = db.TableDef.NewField
'	field.Name = "AGTCOD"
'	field.Description = ""
'	field.Type = WI_CHAR_FIELD
'	field.Equation = ""
'	field.Length = 16
'	task.ReplaceField "AGTCODE", field
'	task.PerformTask	
	
	Set field = db.TableDef.NewField
	field.Name = "CUSNAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 60
	task.ReplaceField "CUSNM", field
	task.PerformTask

'	Set field = db.TableDef.NewField
'	field.Name = "PYMNTTYPE"
'	field.Description = ""
'	field.Type = WI_CHAR_FIELD
'	field.Equation = ""
'	field.Length = 2
'	task.ReplaceField "PYMNTYP", field
'	task.PerformTask

	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
If haveRecords("REPORTSVR.VDPUTR011.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR011.IMD")

	 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "AGTCOD"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	
		Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RECEIPTNO"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 32
	task.ReplaceField "TRNREF", field
	task.PerformTask
	
	Set field = db.TableDef.NewField
	field.Name = "CUSNAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 60
	task.ReplaceField "CUSNM", field
	task.PerformTask
	
	Set field = db.TableDef.NewField
	field.Name = "PYMNTTYPE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 2
	task.ReplaceField "PAYTYP", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
	
	
If haveRecords("REPORTSVR.VDPUTR021.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR021.IMD")

	 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "AGTCOD"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	
		Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARR"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
End Function	

Function EExt_History
If haveRecords("REPORTSVR.VDPUTR04N.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR04N.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "ACCTNO"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "TRNDAT"
	task.AddFieldToInc "POSTDAT"
	task.AddFieldToInc "TRNTYP"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "NUMUNITS"
	task.AddFieldToInc "VALUNITS"
	task.AddFieldToInc "PYMNTTYPE"
	task.AddFieldToInc "CUSNAME"
	task.AddFieldToInc "AGTCOD"
	task.AddFieldToInc "BRANCHCODE"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	task.AddFieldToInc "NARR"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "VALUNITS_NEW"
	dbName = "DPUTR04N.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
If haveRecords("REPORTSVR.VDPURF031.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPURF031.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "ACCTNO"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "TRNDAT"
	task.AddFieldToInc "POSTDAT"
	task.AddFieldToInc "TRNTYP"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "NUMUNITS"
	task.AddFieldToInc "VALUNITS"
	task.AddFieldToInc "PYMNTTYPE"
	task.AddFieldToInc "CUSNAME"
	task.AddFieldToInc "AGTCOD"
	task.AddFieldToInc "BRANCHCODE"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	task.AddFieldToInc "NARR"
	dbName = "DPURF031.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
		
If haveRecords("REPORTSVR.VDPUTR011.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR011.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "ACCTNO"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "TRNDAT"
	task.AddFieldToInc "POSTDAT"
	task.AddFieldToInc "TRNTYP"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "NUMUNITS"
	task.AddFieldToInc "VALUNITS"
	task.AddFieldToInc "PYMNTTYPE"
	task.AddFieldToInc "CUSNAME"
	task.AddFieldToInc "AGTCOD"
	task.AddFieldToInc "BRANCHCODE"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	task.AddFieldToInc "NARR"
	dbName = "DPUTR011.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
If haveRecords("REPORTSVR.VDPUTR021.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR021.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "ACCTNO"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "TRNDAT"
	task.AddFieldToInc "POSTDAT"
	task.AddFieldToInc "TRNTYP"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "NUMUNITS"
	task.AddFieldToInc "VALUNITS"
	task.AddFieldToInc "PYMNTTYPE"
	task.AddFieldToInc "CUSNAME"
	task.AddFieldToInc "AGTCOD"
	task.AddFieldToInc "BRANCHCODE"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	task.AddFieldToInc "NARR"
	dbName = "DPUTR021.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
		
If haveRecords("REPORTSVR.VDPUTR081.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTR081.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "ACCTNO"
	task.AddFieldToInc "TRUSTCOD"
	task.AddFieldToInc "TRNDAT"
	task.AddFieldToInc "POSTDAT"
	task.AddFieldToInc "TRNTYP"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNITPRC"
	task.AddFieldToInc "NUMUNITS"
	task.AddFieldToInc "VALUNITS"
	task.AddFieldToInc "PYMNTTYPE"
	task.AddFieldToInc "CUSNAME"
	task.AddFieldToInc "AGTCOD"
	task.AddFieldToInc "BRANCHCODE"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	task.AddFieldToInc "NARR"
	dbName = "DPUTR081.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

'Append History Files
Function FAppend_History
If haveRecords("DPUTR04N.IMD") Then
Set db = Client.OpenDatabase("DPUTR04N.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "VALUNITS"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "VALUNITS_NEW"
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing

If haveRecords("DPUTR04N.IMD") Then
Set db = Client.OpenDatabase("DPUTR04N.IMD")
	Set task = db.AppendDatabase
   Set pm = Client.ProjectManagement
	If pm.DoesDatabaseExist("DPURF031.IMD") Then
			task.AddDatabase "DPURF031.IMD"
	End If
	If pm.DoesDatabaseExist("DPUTR021.IMD") Then
			task.AddDatabase  "DPUTR021.IMD"
	End If
	If pm.DoesDatabaseExist("DPUTR011.IMD") Then
			task.AddDatabase  "DPUTR011.IMD"
	End If
	If pm.DoesDatabaseExist("DPUTR081.IMD") Then
			task.AddDatabase  "DPUTR081.IMD"
	End If
		dbName = "Master_History.IMD"
   Set pm = Nothing
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

' Join Transactions to Master to get UTCID
Function GJoin_Hist_Master
If haveRecords("Master_History.IMD") Then
Set db = Client.OpenDatabase("Master_History.IMD")
	Set task = db.JoinDatabase
   If haveRecords("Master_Funds.IMD") Then
	task.FileToJoin "Master_Funds.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "UTCID"
	task.AddMatchKey "ACCTNO", "ACCTNOP", "A"
	task.AddMatchKey "TRUSTCOD", "TRUSTCOD", "A"
	task.CreateVirtualDatabase = False
	dbName = "Master_History_File.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function

' Join Transactions to Customers
Function HJoinCustomers
If haveRecords("Master_History_File.IMD")  And haveRecords("REPORTSVR.VDPUCID01.IMD") Then
Set db = Client.OpenDatabase("Master_History_File.IMD")
	Set task = db.JoinDatabase 
	task.FileToJoin "REPORTSVR.VDPUCID01.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "ADDRESS1"
	task.AddSFieldToInc "ADDRESS2"
	task.AddSFieldToInc "ADDRESS3"
	task.AddSFieldToInc "DOB"
	task.AddSFieldToInc "OCCUP"
	task.AddSFieldToInc "HOLDRTYP"
	task.AddSFieldToInc "HOLDRNAME"
	task.AddSFieldToInc "BRANCHCD1"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Transaction_History_INT.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
 
	Set task = Nothing
	Set db = Nothing
End If

If haveRecords("Transaction_History_INT.IMD") Then
	Set db = Client.OpenDatabase("Transaction_History_INT.IMD")
	Set task = db.ExportDatabase
	task.IncludeFieldNames = TRUE
	task.IncludeAllFields
	eqn = ""
	task.Separators "|", "."
	task.PerformTask Client.WorkingDirectory & "Exports.ILB\Transaction_History_INT.DEL", "Database", "DEL", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End If 
If haveRecords("Transaction_History_INT.IMD") Then
	dbName = "Transaction_History_LITE.IMD"
	Client.ImportDelimFile Client.WorkingDirectory & "Exports.ILB\Transaction_History_INT.DEL", dbName, FALSE, "", Client.WorkingDirectory & "Import Definitions.ILB\Transaction_History_INT.RDF", TRUE
End If 

End Function

Function IJoinBranch
If haveRecords("Transaction_History_lite.IMD") And haveRecords("REPORTSVR.VDPBRANCH.IMD") Then
Set db = Client.OpenDatabase("Transaction_History_lite.IMD")
	Set task = db.JoinDatabase
 
	task.FileToJoin "REPORTSVR.VDPBRANCH.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "LONGDES"
	task.AddMatchKey "BRANCHCODE", "BRNCHCODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "Transaction_History.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
 
End Function

' Append Field
Function JCreate_Fields
If haveRecords("Transaction_History.IMD") Then
Set db = Client.OpenDatabase("Transaction_History.IMD")



	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PAYMENT_TYPE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@CompIf((PYMNTTYPE == ""C"" .OR. PYMNTTYPE == ""CASH""),""CASH"", (PYMNTTYPE == ""DD"" .OR. PYMNTTYPE == ""Q""),""CHEQUE"", PYMNTTYPE == ""BATCH"",""STANDING ORDER-BATCH"", PYMNTTYPE == ""D"",""DISTRIBUTION"", PYMNTTYPE == ""B"",""BALANCE"",1,PYMNTTYPE)"
	field.Length = 15
	task.AppendField field
	task.PerformTask

	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_CHANNEL"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@CompIf(PYMNTTYPE == ""WIRE"",""WIRE"", PYMNTTYPE == ""CARD"",""ATM"", RECEIPTNO = ""POS"",""POINT OF SALE"",1,"""")"
	field.Length = 15
	task.AppendField field
	task.PerformTask



'	Set task = db.TableManagement
'	Set field = db.TableDef.NewField
'	field.Name = "PAYMENT_TYPE"
''	field.Description = ""
'	field.Type = WI_VIRT_CHAR
'	field.Equation = "@CompIf(PYMNTTYPE == ""C"",""CASH"", PYMNTTYPE == ""Q"",""CHEQUE"", PYMNTTYPE == ""S"",""STANDING ORDER"", PYMNTTYPE == ""D"",""DISTRIBUTION"", PYMNTTYPE == ""B"",""BALANCE"",1,"""")"
'	field.Length = 15
'	task.AppendField field
'	task.PerformTask

'	Set field = db.TableDef.NewField
'	field.Name = "TRANSACTION_CHANNEL"
''	field.Description = ""
'	field.Type = WI_VIRT_CHAR
'	field.Equation = "@CompIf(RECEIPTNO = ""WT-"",""WIRE"", RECEIPTNO = ""VATM"",""ATM"", RECEIPTNO = ""POS"",""POINT OF SALE"",1,"""")"
'	field.Length = 15
'	task.AppendField field
'	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_TYPE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@CompIf(TRNTYP == ""R"",""REPO"", TRNTYP== ""S"",""SALE"",1,"""")"
	field.Length = 10
	task.AppendField field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "HOLDER_TYPE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@CompIf(@left(HOLDRTYP,1) == ""C"",""CORPORATION"", @Left(HOLDRTYP,1) == ""I"",""INDIVIDUAL"",1,"""")"
	field.Length = 15
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

' Rename Fields
Function KRename_SourceFields
If haveRecords("Transaction_History.IMD") Then
Set db = Client.OpenDatabase("Transaction_History.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BRANCH_DESCRIPTION"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 55
	task.ReplaceField "LONGDES", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_BRANCH"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 20
	task.ReplaceField "BRANCHCODE", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_BRANCH"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 20
	task.ReplaceField "BRANCHCD1", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "AGENT_CODE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 10
	task.ReplaceField "AGTCOD", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "UNIT_PRICE"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "UNITPRC", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "UNITS"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "NUMUNITS", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NARRATIVE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 40
	task.ReplaceField "NARR", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DATE_OF_BIRTH"
	field.Description = ""
	field.Type = WI_DATE_FIELD
	field.Equation = "YYYYMMDD"
	task.ReplaceField "DOB", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "HOLDER_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 60
	task.ReplaceField "HOLDRNAME", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ACCT_NO"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 15
	task.ReplaceField "ACCTNO", field
	task.PerformTask
		
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CUSTOMER_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 50
	task.ReplaceField "CUSNAME", field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRUST_CODE"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 3
	task.ReplaceField "TRUSTCOD", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRANSACTION_AMOUNT"
	field.Description = ""
	field.Type = WI_NUM_FIELD
	field.Equation = ""
	field.Decimals = 2
	task.ReplaceField "VALUNITS", field
	task.PerformTask
	

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "OCCUPATION"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "OCCUP"
	field.Length = 30
	task.AppendField field
	task.PerformTask	 
		
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function


'Create Date Fields
Function LCreate_DateFields
If haveRecords("Transaction_History.IMD") Then
Set db = Client.OpenDatabase("Transaction_History.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "POST_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(POSTDAT, ""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TRAN_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(TRNDAT, ""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask

	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function


'Amended April 17, 2015 to rename output file to make the changes below.
Function MExt_Physical
If haveRecords("Transaction_History.IMD") Then
Set db = Client.OpenDatabase("Transaction_History.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "UTCID"
	task.AddFieldToInc "CUSTOMER_BRANCH"
	task.AddFieldToInc "ACCT_NO"
	task.AddFieldToInc "TRUST_CODE"
	task.AddFieldToInc "PAYMENT_TYPE"
	task.AddFieldToInc "TRANSACTION_CHANNEL"
	task.AddFieldToInc "TRANSACTION_BRANCH"
	task.AddFieldToInc "BRANCH_DESCRIPTION"
	task.AddFieldToInc "TRANSACTION_TYPE"
	task.AddFieldToInc "HOLDER_TYPE"
	task.AddFieldToInc "POST_DATE"
	task.AddFieldToInc "TRAN_DATE"
	task.AddFieldToInc "UNITS"
	task.AddFieldToInc "RECEIPTNO"
	task.AddFieldToInc "UNIT_PRICE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "CUSTOMER_NAME"
	task.AddFieldToInc "AGENT_CODE"
	task.AddFieldToInc "NARRATIVE"
	task.AddFieldToInc "HOLDER_NAME"
	task.AddFieldToInc "ADDRESS1"
	task.AddFieldToInc "ADDRESS2"
	task.AddFieldToInc "ADDRESS3"
	task.AddFieldToInc "DATE_OF_BIRTH"
	task.AddFieldToInc "OCCUPATION"
	task.AddFieldToInc "TRANSRC"
	task.AddFieldToInc "TRANSTATUS"
	dbName = "History_Transaction_Hist.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

'Created April 17 to create Risk Ratings for all customers to include in all Results
Function MFJoinPortaltoRisk
   If haveRecords("dbo.EvaluationMatrix.IMD") Then
	Set db = Client.OpenDatabase("dbo.EvaluationMatrix.IMD")
	Set task = db.JoinDatabase
   If haveRecords("Risk_Ratings-" & Trim(e_Risk_Rating_Sheet_Name) & ".IMD") Then	
	task.FileToJoin ("Risk_Ratings-" & Trim(e_Risk_Rating_Sheet_Name) & ".IMD")
	task.AddPFieldToInc "UTCID"
	task.AddPFieldToInc "PART2_8"
	task.AddSFieldToInc "UTCID"
	task.AddSFieldToInc "NAME"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.CreateVirtualDatabase = False
	dbName = "Risk Ratings and PORTAL.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_REC
End If
	Set task = Nothing
	Set db = Nothing
End If
	Client.CloseAll
End Function

'Created April 17, 2015 to get Customer Branch Name
Function MG_CreateKeys
If haveRecords("RISK RATINGS.IMD") Then	
 	Set db = Client.OpenDatabase("RISK RATINGS.IMD") 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RATING_SOURCE"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """CORE"""
	field.Length = 10
	task.AppendField field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RISK_RATING"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@ALLTRIM(RISKRATING)"
	field.Length = 10
	task.AppendField field
	task.PerformTask
 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "NEW_UTCID"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@ALLTRIM(UTCID)"
	field.Length = 10
	task.AppendField field
	task.PerformTask	
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
End Function

'Created April 17, 2015 to get Risk Rating and Source for each customer
Function MHJoin_Portal
If haveRecords("History_Transaction_Hist.IMD") And haveRecords("RISK RATINGS.IMD") Then
Set db = Client.OpenDatabase("History_Transaction_Hist.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "RISK RATINGS.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "RATING_SOURCE"
	task.AddSFieldToInc "RISK_RATING"
	task.AddMatchKey "UTCID", "UTCID", "A"
	task.CreateVirtualDatabase = False
	dbName = "History_Transaction_Hist_Risk.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
 
End Function


'Created April 17, 2015 to get Customer Branch Name
Function MI_JoinCustomerBranch
If haveRecords("History_Transaction_Hist_Risk.IMD") Then
Set db = Client.OpenDatabase("History_Transaction_Hist_Risk.IMD")
	Set task = db.JoinDatabase
   If haveRecords("REPORTSVR.VDPBRANCH.IMD") Then
	task.FileToJoin "REPORTSVR.VDPBRANCH.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "LONGDES"
	task.AddMatchKey "CUSTOMER_BRANCH", "BRNCHCODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "History_Transaction_History.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
End If
	Set task = Nothing
	Set db = Nothing
End If
End Function

'Amended April 17, 2015 to include Customer Branch Name
'Extract Daily Transactions
Function NExtHist_Daily
If haveRecords("History_Transaction_History.IMD") Then
Set db = Client.OpenDatabase("History_Transaction_History.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "RECEIPT_NUMBER"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 20
	task.ReplaceField "RECEIPTNO", field
	task.PerformTask
	
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BRANCH_NAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 15
	task.ReplaceField "LONGDES", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
If haveRecords("History_Transaction_History.IMD") Then
Set db = Client.OpenDatabase("History_Transaction_History.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Daily_Transactions_Today.IMD"
	task.AddExtraction dbName, "", "@Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE"" .AND. POST_DATE == "& CHR(34) & v_currentDateChar & CHR(34)
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

'Extract History Transactions for Average
Function OExtHist_Average
If haveRecords("History_Transaction_History.IMD") Then
Set db = Client.OpenDatabase("History_Transaction_History.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Tran_Hist_Average.IMD"
	task.AddExtraction dbName, "", "@Match(TRANSACTION_TYPE, ""SALE"", ""REPO"") .AND. PAYMENT_TYPE <> ""BALANCE""  .AND. @BetweenDate(POST_DATE, "& CHR(34) &  v_AverageDateStartChar & CHR(34) &", "& CHR(34) &  v_AverageDateEndChar & CHR(34) & ")"	
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

'Extract Daily New Accounts
Function QCreateAccountDate
If haveRecords("REPORTSVR.VCRDHLDR.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VCRDHLDR.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "STATUS"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@If(CARDIND==""Y"", ""ACTIVE"",""INACTIVE"")"
	field.Length = 15
	task.AppendField field
	task.PerformTask
	
	 
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CHGDATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "STARTDATE_DATE"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
 
End If

If haveRecords("REPORTSVR.VDPUTP01.IMD") Then
Set db = Client.OpenDatabase("REPORTSVR.VDPUTP01.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CREATE_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(@left(@alltrim(RECDATE),8), ""YYYYMMDD"")"
'	field.Equation = "RECDATE1"
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End If
	Set field = Nothing
	
	

If haveRecords("REPORTSVR.VPAYEE.IMD") Then	
		Set db = Client.OpenDatabase("REPORTSVR.VPAYEE.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PAYEENAME"
	field.Description = ""
	field.Type = WI_CHAR_FIELD
	field.Equation = ""
	field.Length = 560
	task.ReplaceField "PAYEE_NAME", field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
end if
End Function

Function RExtAccounts_Daily
If haveRecords("REPORTSVR.VDPUTP01.IMD") Then
	Set db = Client.OpenDatabase("REPORTSVR.VDPUTP01.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Accounts_Created_Today.IMD"
	task.AddExtraction dbName, "", "CREATE_DATE ==  "& Chr(34) & v_currentDateChar & Chr(34)
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function

Function SCleanup
	DeleteFile("Master_Funds.IMD") 
	DeleteFile("Master_History.IMD")
	DeleteFile("DPUTR04N.IMD")
	DeleteFile("DPURF031.IMD")
	DeleteFile("DPUTR011.IMD")
	DeleteFile("DPUTR021.IMD")
	DeleteFile("DPUTR081.IMD")
	DeleteFile("Master_History_File.IMD")
	DeleteFile("Transaction_History_INT.IMD")
	DeleteFile("Transaction_History.IMD")
	DeleteFile("History_Transaction_Hist.IMD")
	DeleteFile("Risk_Ratings-" & Trim(e_Risk_Rating_Sheet_Name) & ".IMD")
	DeleteFile("Risk Ratings and PORTAL.IMD")
	DeleteFile("History_Transaction_Hist_Risk.IMD")	
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




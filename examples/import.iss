'====================================================================================================
'	Test#: 		Import
'	Risk:		
' 	Objective:	
' 	Frequency:	
' 	Last Modified:	12/28/2014 3:55:12 PM
'====================================================================================================
'	Script Dependencies:
'====================================================================================================

'----- Constants -----

Const RESULT_FILENAME = "Import"
Const scriptname_log ="Import.iss"
Global errors_string As String
Const division = "UTC"
Dim fso As Object
Dim Echo As String
Dim Echo2 As String

Sub Main

 	Ignorewarning(True)
	Set fso    = CreateObject("Scripting.FileSystemObject")
 	On Error GoTo finalRoutine

	


'	LogMessage "Formatting dates"
	v_Import_From =  Format(CDate(CLng(Date) - 2) , "dd-mm-yyyy")        ' to be used at go-live to import data for the previous day
	v_Import_To =  Format(CDate(CLng(Date) - 2) , "dd-mm-yyyy")        ' to be used at go-live to import data for the previous day
	v_Import_From2 = Format(CDate(CLng(Date) - e_Watch_Reactivated_Accounts_Count) , "dd-mm-yyyy")       ' Get a larger section of file to do test 
	


'e_Echo = "si05prf"
Echo = "caseware" 	' WMS PASSWORD
Echo2 = "caseware2023"  ' CORE PASSWORD
'MsgBox Echo
 


'	Call ACleanup()
	GET_Items
 If e_DR_CONNECTION  Then
 
'	 MsgBox("DR")
	DR_ODBCImport


Else 

 '	MsgBox("LIVE")
	Live_ODBCImport
End If

	CImportRiskRating
	H_ImportAgentCodes
	DImportInternalWatchList
	G_ImportBlacklist


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



Function Live_ODBCImport


	dbName = "RISK RATINGS.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "CASEWAREBENEFICIARYVIEW" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo2 & ";DBQ=COREPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", "select  ORGKEY AS " & Chr(34) & "UTCID" & Chr(34) & ", riskrating from CUSTOM.CASEWARECUSTOMERVIEW  UNION ALL  Select ORGKEY As " & Chr(34) & "UTCID" & Chr(34) & ",riskrating from CUSTOM.CASEWARECORPCUSTOMERVIEW"
			Count_Database_Records(dbName)
	Client.CloseAll
	
	dbName = "REPORTSVR.VDPUTR011.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR011" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPURF031.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPURF031" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VPAYEE.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VPAYEE" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	
	dbName = "REPORTSVR.VDPUTR017.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR017" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR04N.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR04N" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
 
	dbName = "REPORTSVR.VDPUCID01_int.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "UTCID" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo2 & ";DBQ=COREPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll	
 
	Client.CloseAll
	dbName = "dbo.Registrations_int.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "REGISTRATIONS" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo2 & ";DBQ=COREPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR021.IMD" 
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR021" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTP01.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTP01" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR081.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR081" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPBRANCH.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPBRANCH" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	
	dbName = "REPORTSVR.VCRDHLDR.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VCRDHLDR" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll

'	dbName = "REPORTSVR.VDPFEX.IMD"
'	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPFEX" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
'		Count_Database_Records(dbName)
'	Client.CloseAll

'	dbName = "RISK RATINGS.IMD"
'	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "CASEWAREBENEFICIARYVIEW" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo2 & ";DBQ=COREPROD;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", "select  ORGKEY AS " & Chr(34) & "UTCID" & Chr(34) & ", riskrating from CUSTOM.CASEWARECUSTOMERVIEW  UNION ALL  Select ORGKEY As " & Chr(34) & "UTCID" & Chr(34) & ",riskrating from CUSTOM.CASEWARECORPCUSTOMERVIEW"
'			Count_Database_Records(dbName)
'	Client.CloseAll
End Function


' File - Import Assistant: Excel Risk Ratings
Function CImportRiskRating
RiskRating = Dir( Trim(e_Compliance_Input_Area) & Trim(e_Risk_Rating_File_Name) & "*.xlsx")
		If RiskRating <> "" Then
		Set task = Client.GetImportTask("ImportExcel")
			dbName = Trim(e_Compliance_Input_Area) & RiskRating
			task.FileToImport = dbName
			task.SheetToImport = Trim(e_Risk_Rating_Sheet_Name)
			task.OutputFilePrefix = "Risk_Ratings"
			task.FirstRowIsFieldName = "TRUE"
			task.EmptyNumericFieldAsZero = "FALSE"
			task.PerformTask
			dbName = task.OutputFilePath (Trim (e_Risk_Rating_Sheet_Name))
		Set task = Nothing
		Client.CloseAll
	End If 	
End Function

' File - Import Assistant: Excel Risk Ratings
Function DImportInternalWatchList
InternalList = Dir( Trim(e_Compliance_Input_Area) & Trim(e_Internal_List_File_Name) & "*.xlsx")
		If InternalList <> "" Then
		Set task = Client.GetImportTask("ImportExcel")
			dbName = Trim(e_Compliance_Input_Area) & InternalList 
			task.FileToImport = dbName
			task.SheetToImport = Trim(e_Internal_List_Sheet_Name)
			'task.OutputFilePrefix = "AML Compliance Internal Monitoring List"
			task.OutputFilePrefix = "AML Compliance Internal WATCHLIST"
			task.FirstRowIsFieldName = "TRUE"
			task.EmptyNumericFieldAsZero = "FALSE"
			task.PerformTask
			dbName = task.OutputFilePath  (Trim (e_Internal_List_Sheet_Name))
		Set task = Nothing
		Client.CloseAll
	End If 	
End Function

' File - Import Assistant: Excel Risk Ratings
Function F_ImportChurches
InternalList = Dir( Trim(e_Compliance_Input_Area) & Trim(e_Church_File_Name) & "*.xlsx")
		If InternalList <> "" Then
		Set task = Client.GetImportTask("ImportExcel")
			dbName = Trim(e_Compliance_Input_Area) & InternalList 
			task.FileToImport = dbName
			task.SheetToImport = Trim(e_Church_Sheet_Name)
			task.OutputFilePrefix = "Churches, NGOs and Charities"
			task.FirstRowIsFieldName = "TRUE"
			task.EmptyNumericFieldAsZero = "FALSE"
			task.PerformTask
			dbName = task.OutputFilePath  (Trim (e_Church_Sheet_Name))
		Set task = Nothing
		Client.CloseAll
	End If 	
End Function


' File - Import Assistant: Blacklist
Function G_ImportBlacklist
InternalList = Dir( Trim(e_Compliance_Input_Area) & Trim(e_Blacklist_File_Name) & "*.xlsx")
		If InternalList <> "" Then
		Set task = Client.GetImportTask("ImportExcel")
			dbName = Trim(e_Compliance_Input_Area) & InternalList 
			task.FileToImport = dbName
			task.SheetToImport = Trim(e_Blacklist_Sheet_Name)
			task.OutputFilePrefix = "Compliance Blacklist"
			task.FirstRowIsFieldName = "TRUE"
			task.EmptyNumericFieldAsZero = "FALSE"
			task.PerformTask
			dbName = task.OutputFilePath  (Trim (e_Blacklist_Sheet_Name))
		Set task = Nothing
		Client.CloseAll
	End If 	
End Function

Function H_ImportAgentCodes

InternalList = Dir( Trim(e_Compliance_Input_Area) &  "OnlineAgents*.xlsx")
If InternalList <> "" Then
	Set task = Client.GetImportTask("ImportExcel")
 	dbName = Trim(e_Compliance_Input_Area) & InternalList 
	task.FileToImport = dbName
	task.SheetToImport = "Codes"
	task.OutputFilePrefix = "OnlineAgents"
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "FALSE"
	task.PerformTask
	dbName = task.OutputFilePath("Codes")
	Set task = Nothing
End If

If haveRecords("OnlineAgents-Codes.IMD") Then
	Set db = Client.OpenDatabase("OnlineAgents-Codes.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Online_Agents.IMD"
	task.AddExtraction dbName, "", "ONLINE_AGENTS <> """""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
End If
End Function



Function GET_Items

	dbName = "UserRoleAssignments.IMD"
	Client.ImportODBCFile "" & Chr(34) & "dbo" & Chr(34) & "." & Chr(34) & "UserRoleAssignments" & Chr(34) & "", dbName, FALSE, ";DSN=CasewareMOn;UID=;Trusted_Connection=Yes;APP=IDEA;WSID=COBALT;DATABASE=ALESSADB", ""
			Count_Database_Records(dbName)
	
		dbName = "Roles.IMD"
	Client.ImportODBCFile "" & Chr(34) & "dbo" & Chr(34) & "." & Chr(34) & "UserRoles" & Chr(34) & "", dbName, FALSE, ";DSN=CasewareMOn;UID=;Trusted_Connection=Yes;APP=IDEA;WSID=COBALT;DATABASE=ALESSADB", ""
			Count_Database_Records(dbName)
	
		dbName = "Users.IMD"
	Client.ImportODBCFile "" & Chr(34) & "dbo" & Chr(34) & "." & Chr(34) & "Users" & Chr(34) & "", dbName, FALSE, ";DSN=CasewareMOn;UID=;Trusted_Connection=Yes;APP=IDEA;WSID=COBALT;DATABASE=ALESSADB", ""
 			Count_Database_Records(dbName)
End Function


Function DR_ODBCImport
	dbName = "REPORTSVR.VDPUTR011.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR011" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
 		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPURF031.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPURF031" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VPAYEE.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VPAYEE" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	
	dbName = "REPORTSVR.VDPUTR017.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR017" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR04N.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR04N" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
 
	dbName = "REPORTSVR.VDPUCID01_int.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "UTCID" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & "002;DBQ=COREDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "dbo.Registrations_int.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "REGISTRATIONS" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & "002;DBQ=COREDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR021.IMD" 
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR021" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTP01.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTP01" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPUTR081.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPUTR081" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	dbName = "REPORTSVR.VDPBRANCH.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPBRANCH" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll
	
	dbName = "REPORTSVR.VCRDHLDR.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VCRDHLDR" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll

'	dbName = "REPORTSVR.VDPFEX.IMD"
'	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "VDPFEX" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & ";DBQ=WMSDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", ""
		Count_Database_Records(dbName)
	Client.CloseAll

	dbName = "RISK RATINGS.IMD"
	Client.ImportODBCFile "" & Chr(34) & "CUSTOM" & Chr(34) & "." & Chr(34) & "CASEWAREBENEFICIARYVIEW" & Chr(34) & "", dbName, FALSE, ";DSN=FINDEV;UID=caseware;PWD=" & echo & "002;DBQ=COREDR;DBA=W;APA=T;EXC=F;FEN=T;QTO=F;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F;FWC=F;FBS=64000;TLO=O;MLD=0;ODA=F;STE=F;TSZ=8192;AST=FLOAT;", "select  ORGKEY AS " & Chr(34) & "UTCID" & Chr(34) & ", riskrating from CUSTOM.CASEWARECUSTOMERVIEW  UNION ALL  Select ORGKEY As " & Chr(34) & "UTCID" & Chr(34) & ",riskrating from CUSTOM.CASEWARECORPCUSTOMERVIEW"
			Count_Database_Records(dbName)
	Client.CloseAll
End Function



'Cleanup Source and Interim files prior to Importing data
Function ACleanUp
	DeleteFile("Accounts_Created_Today.IMD")
	DeleteFile("Daily_Transactions_Today.IMD")
	DeleteFile("Tran_Hist_Average.IMD")
	DeleteFile("History_Transaction_History.IMD")
'	DeleteFile("Registrations.IMD")	
	DeleteFile("dbo.EvaluationMatrix.IMD")	
	DeleteFile("REPORTSVR.VDPUCID01.IMD")
	DeleteFile("REPORTSVR.VDPUTR021.IMD")
	DeleteFile("REPORTSVR.VDPURF031.IMD")
	DeleteFile("REPORTSVR.VDPUTR04N.IMD")
	DeleteFile("REPORTSVR.VDPUTR011.IMD")
	DeleteFile("REPORTSVR.VDPUTR081.IMD")
	DeleteFile("REPORTSVR.VDPBRANCH.IMD")	
	DeleteFile("REPORTSVR.VDPUTP01.IMD")		
	DeleteFile("REPORTSVR.VCRDHLDR.IMD")
	DeleteFile("REPORTSVR.VPAYEE.IMD")
	DeleteFile("REPORTSVR.VDPBEN.IMD")	
	DeleteFile("REPORTSVR.VDPFEX.IMD")
	DeleteFile("dbo.Registrations.IMD")
'	DeleteFile("Risk_Ratings-" & Trim(e_Risk_Rating_Sheet_Name) & ".IMD")
'	DeleteFile("AML Compliance Internal Monitoring List-Sheet1.IMD")
'	DeleteFile("Compliance Blacklist-Sheet1.IMD")	
'	DeleteFile("Churches, NGOs and Charities-Churches.IMD")	
'	DeleteFile("Risk Ratings and PORTAL.IMD")
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

Function Count_Database_Records(FileName As String)

On Error GoTo err_handle

Dim LogCountFile As TextStream
Dim LogName As String
Dim Path As String

Set db = Client.OpenDatabase(FileName)
recnum = db.count
Path = Client.WorkingDirectory
LogName = "LogCounter_" &  Format(Date, "YYYYMMMDD") & ".txt"

'Create the log if it does not exist and writes reader record
  If Not fso.FileExists(Path & LogName) Then
 
	fso.CreateTextFile (Path & LogName) 
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 2, True)
	LogCountFile.WriteLine("Log_Date" & Chr(9) & Chr(9) & Chr(9) & "Database_Name" & Chr(9) & Chr(9) & Chr(9) & "Record_Count")
	LogCountFile.Close
 
  End If 

' Writes records To file that already exists
	Set LogCountFile = fso.OpenTextFile(Path & LogName, 8, True) 
	LogCountFile.WriteLine (Now & Chr(9) & Chr(9) & Chr(9) & FileName & Chr(9) & Chr(9) & Chr(9) & recnum)
	LogCountFile.Close
	Set db = Nothing
err_handle:
    Client.CloseAll
End Function

Function DeleteFile(NameOfFile As String)
	Client.CloseAll
	If fso.FileExists(Client.WorkingDirectory & Trim(NameOfFile)) = True Then Kill(Client.WorkingDirectory & Trim(NameOfFile))
End Function



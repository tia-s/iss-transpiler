'Cut Customer Fles to get created up to 2 days ago
Function C_Get_Customers
If haveRecords("NameScan_Batch.IMD") Then
Set db = Client.OpenDatabase("NameScan_Batch.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CREATE_DATE"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ntod(RECDATE1 -1000000, ""YYMMDD"")"
	task.AppendField field
	task.PerformTask
	Set field = Nothing
End If	
	Set task = Nothing
	Set db = Nothing
End Function

'Extract fields from Customers
Function D1_Create_Comp_DOB
If haveRecords("NameScan_All.IMD") Then
	Set db = Client.OpenDatabase("NameScan_All.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "COMP_DOB"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@If(BIRTHDAT > 1000000, @Val(@Mid(@Str(BIRTHDAT, 7, 0), 2, 6)), BIRTHDAT)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "BIRTH"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@if(BIRTHDAT >= 1000000, ""20""+@Mid(@Str(BIRTHDAT, 7, 0) , 2, 6), ""19"" + @Mid(@Str(BIRTHDAT, 6, 0), 1, 6)) "
	field.Length = 8
	task.AppendField field
	task.PerformTask
	
End If	
	Set task = Nothing
	Set db = Nothing
End Function		

Function D2_Create_DOB
If haveRecords("NameScan_All.IMD") Then
	Set db = Client.OpenDatabase("NameScan_All.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DOB"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@ctod(@compif(COMP_DOB < 10101, ""0"", COMP_DOB > 991231, ""20"" + @Str(COMP_DOB, 6, 0), @between(COMP_DOB, 150732, 991231), ""19"" + @Str(COMP_DOB, 6, 0), 1, ""20"" + @Str(COMP_DOB, 6, 0) ), ""YYYYMMDD"")"
	field.Equation = "@ctod(@compif(COMP_DOB < 10101, ""0"", COMP_DOB > 991231, ""20"" + @Str(COMP_DOB, 6, 0), @between(COMP_DOB, @Val(@Mid(Chr(34) & v_currentDateChar  & Chr(34), 4, 6)), 991231), ""19"" + @Str(COMP_DOB, 6, 0), 1, ""20"" + @Str(COMP_DOB, 6, 0) ), ""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask

	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DOB"
	field.Description = ""
	field.Type = WI_VIRT_DATE
	field.Equation = "@Ctod(BIRTH, ""YYYYMMDD"")"
	task.AppendField field
	task.PerformTask
End If	
	Set task = Nothing
	Set db = Nothing
End Function


' Modify Source Fields to Match screening requirements
Function H_Create_Key_Fields
If haveRecords("Customer.IMD") Then
	Set db = Client.OpenDatabase("Customer.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
		field.Name = "REQUEST_ID"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@ALLTRIM(UTCID)"
		field.Length = 25
		task.AppendField field
		
	Set field = db.TableDef.NewField
		field.Name = "ADDRESS"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@ALLTRIM(ADDRESS1) +"" "" + @ALLTRIM(ADDRESS2) + "" "" + @ALLTRIM(ADDRESS3) + "" "" +  @ALLTRIM(COUNTRY)"
		field.Length = 100
		task.AppendField field

	Set field = db.TableDef.NewField
		field.Name = "INDIVIDUAL_ORGANIZATION"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@CompIf( HOLDRTYP  == ""C"", ""ORGANIZATION"",1, ""PERSON"" )"
		field.Length = 25
		task.AppendField field

	Set field = db.TableDef.NewField
		field.Name = "TAX_ID"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation =  "@if(ID<> """", @ALLTRIM(ID), @ALLTRIM(PP))"
		field.Length = 25
		task.AppendField field

	Set field = db.TableDef.NewField
		field.Name = "NATIONALITY"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = "COUNTRY"
		field.Length = 50
		task.AppendField field

		Set field = db.TableDef.NewField
		field.Name = "ALIAS"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@ALLTRIM(OTHERNAMES)"
		field.Length = 20
		task.AppendField field

		Set field = db.TableDef.NewField
		field.Name = "POSTCODE"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """"""
		field.Length = 20
		task.AppendField field
		
		Set field = db.TableDef.NewField
		field.Name = "NAME"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@ALLTRIM(FIRSTNAME) +"" "" + @ALLTRIM(SURNAME)"
		field.Length = 50
		task.AppendField field

		Set field = db.TableDef.NewField
		field.Name = "GENDER"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = "SEX"
		field.Length = 1
		task.AppendField field
		
		
		Set field = db.TableDef.NewField
		field.Name = "BUSINESS_UNIT"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = ""UTCBATCH""
		field.Length = 15
		task.AppendField field
		
		Set field = db.TableDef.NewField
		field.Name = "PAYLOAD_DATA"
		field.Description = ""
		field.Type = WI_CHAR_FIELD
		field.Equation = """"
		field.Length = 15
		task.AppendField field
		
		Set task = db.TableManagement
		Set field = db.TableDef.NewField
		field.Name = "VESSEL_IDENTIFICATION"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = """"""
		field.Length = 2
		task.AppendField field
		task.PerformTask		
	
		Set field = db.TableDef.NewField
		field.Name = "DATEOFBIRTH"
		field.Description = ""
		field.Type = WI_VIRT_CHAR
		field.Equation = "@Dtoc(DOB,""YYYY/MM/DD"")"
		field.Length = 15
		task.AppendField field
		task.PerformTask

		
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "CATEGORY"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """"""
	field.Length = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
		
		
		
		
		
End If
End Function




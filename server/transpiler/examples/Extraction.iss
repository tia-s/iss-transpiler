Set db = Client.OpenDatabase("DB1.IMD") ' check if this is an idea keyword
Set task = db.Extraction ' start of extraction
task.IncludeAllFields
dbName = "DB5.IMD"
task.AddExtraction dbName, "", ""
task.CreateVirtualDatabase = False
task.PerformTask 1, db.Count ' end of extraction
Set task = Nothing
Set db = Nothing

Set task = db.Extraction
task.AddFieldToInc  "COL1"
task.AddFieldToInc  "COL2"
task.AddFieldToInc "COL3"
dbName = "DB4.IMD"
task.AddExtraction dbName, "",  "@AllTrim(@Strip(COL1)) <> """" .OR. @AllTrim(@Strip(COL2)) <> """""
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing

Set db = Client.OpenDatabase("DB2.IMD")
Set task = db.Extraction
task.AddFieldToInc "COl1"
task.AddFieldToInc "COL2"
task.AddFieldToInc "COL3"
task.AddKey "COL3", "A"
dbName = "DB3.IMD"
task.AddExtraction dbName, "", ""
task.CreateVirtualDatabase = False
task.PerformTask 1, db.Count
Set task = Nothing
Set db = Nothing
Client.OpenDatabase(dbName)

' (separate from extraction)... optional set db

' set task
' don't constrain order for below:
'''''''
' add field to include | include all fields
' dbname 
' add extraction task with filter
' optional create virtual db
' optional add key
'''''''''
' perform task with options for number of records to extract

' (separate from extraction)... set tasks to nothing
' (separate from extraction)... optional client open database

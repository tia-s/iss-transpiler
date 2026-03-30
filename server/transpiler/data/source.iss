' Example
Set task = db.Extraction ' start of extraction
Set task = Nothing
Set db = Nothing
Set db = Client.OpenDatabase("DB1.IMD") ' open_db
Set task = db.Extraction ' start of extraction
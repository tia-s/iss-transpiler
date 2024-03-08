import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')

# Use date fields to get current date for new customers extract.
v_currentDateChar = ''
"""
v_CurrentDate = CDate(CLng(Date) - 1)
v_CurrentDateChar = Format(v_CurrentDate, "YYYYMMDD")		
"""

# Import Customer Data
def get_data():
    staging = {
        'DSN': 'AS400PROD',
        'TABLES':{'SELECT * FROM REPORTSVR.VDPUCID01': 'NameScan_Batch',
                    'SELECT * FROM REPORTSVR.VDPUTP01': 'REPORTSVR_VDPUTP01'},
    }

    cnxn = pypyodbc.connect(f"DSN={staging['DSN']};")

    for query, table_name in staging["TABLES"].items():
        wd.importSQL(cnxn, query=query, tblName=table_name)
        print("%s: Import Complete." % table_name)

# Cut Customer Fles to get created up to 2 days ago
def get_customers():
    if not wd.db("NameScan_Batch").empty:
        wd.open("NameScan_Batch")
        wd.addCol("CREATE_DATE", lambda row: dt.datetime.strptime(str(int(row.RECDATE1) - 1000000), "%Y%m%d"))
        wd.extract("NameScan_All", filter=f"CREATE_DATE < {v_currentDateChar}")

# Extract fields from Customers
def create_comp_dob():
    if not wd.db("NameScan_All").empty:
        wd.open("NameScan_All")
        wd.addCol("BIRTH", lambda row: '20'+ str(row.BIRTHDAT)[1:7] if row.BIRTHDAT >= 1000000 else '19' + str(row.BIRTHDAT)[0:6])

def create_dob():
    if not wd.db("NameScan_All").empty:
        wd.open("NameScan_All")
        wd.addCol("DOB", lambda row: dt.datetime.strptime(row.BIRTH, "%Y%m%d"))

# Join with Products to only get Customers with Accounts.
def join_accounts():
    if not wd.db("NameScan_All").empty:
        wd.open("NameScan_All")
        wd.join("NameScan_Accounts1", right=wd.db("REPORTSVR_VDPUTP01")[["UTCID", "ACCTNOP"]], how="inner", on=["UTCID"])
        wd.extract("NameScan_Accounts", filter="SURNAME != 'closed'")

# Export source fields to screen
def extract_customers():
    if not wd.db("NameScan_Accounts").empty:
        wd.open("NameScan_Accounts")
        wd.extract("Customer", cols=["UTCID","ID","PP","SURNAME","HOLDRTYP","FIRSTNAME","DOB","OTHERNAMES","ADDRESS1","ADDRESS2","ADDRESS3","COUNTRY","SEX","OCCUP"])

# Modify Source Fields to Match screening requirements
def create_key_fields():
    if not wd.db("Customer").empty:
        wd.open("Customer")
        wd.addCol("REQUEST_ID", lambda row: row.UTCID.strip())
        wd.addCol("ADDRESS", lambda row: f"{row.ADDRESS1.strip()} {row.ADDRESS2.strip()} {row.ADDRESS3.strip()} {row.COUNTRY.strip()}" )
        wd.addCol("INDIVIDUAL_ORGANIZATION", lambda row: 'ORGANIZATION' if row.HOLDRTYP == 'C' else 'PERSON')
        wd.addCol("TAX_ID", lambda row: row.ID.strip() if row.ID != '' else row.PP.strip())
        wd.addCol("NATIONALITY", lambda _: "COUNTRY")
        wd.addCol("ALIAS", lambda row: row.OTHERNAMES.strip())
        wd.addCol("POSTCODE", lambda _: '')
        wd.addCol("NAME", lambda row: f"{row.FIRSTNAME.strip()} {row.SURNAME.strip()}")
        wd.addCol("GENDER", lambda _: "SEX")
        wd.addCol("BUSINESS_UNIT", lambda _: "UTCBATCH")
        wd.addCol("PAYLOAD_DATA", lambda _: '')
        wd.addCol("VESSEL_IDENTIFICATION", lambda _: '')
        wd.addCol("DATEOFBIRTH", lambda row: dt.datetime.strptime(row.DOB, "%Y/%m/%d"))
        wd.addCol("CATEGORY", lambda _: '')

# Export to CSV for Auto Screening Job to start
def export_database_csv():
    if not wd.db("Customer").empty:
        wd.open("Customer")
        wd.extract("Customer_Export", cols=["REQUEST_ID","TAX_ID","NAME","ALIAS","ADDRESS","POSTCODE","COUNTRY","DATEOFBIRTH","GENDER","NATIONALITY","INDIVIDUAL_ORGANIZATION"])
        wd.exportFile(format='csv', sep=",", filename=join(PARENT_DIR, 'Red point', 'Input', 'Screening_Extract'))

def cleanup():
    if wd.tblName:
        wd.close()
    wd.delete("NameScan_Batch")
    wd.delete("NameScan_All")
    wd.delete("NameScan_Accounts")
    wd.delete("REPORTSVR_VDPUTP01")

if __name__ == "__main__":
    wd = DataAnalytics()
    
    # Delete RedPoint Source File prior to new Run		
    
    get_data()
    get_customers()
    create_comp_dob()
    create_dob()
    join_accounts()
    extract_customers()
    create_key_fields()
    export_database_csv()
    cleanup()


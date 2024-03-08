import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')

v_YearMonth = ''
"""
' Use date fields to get current Month start and end date.
	v_CurrentDate = CDate(CLng(Date) - 1)
		
'	v_CurrentDate = "2022-03-01"
	v_CurrentDateChar = Format(v_CurrentDate, "YYYYMMDD")

	v_YearMonth = Mid(v_CurrentDateChar, 1, 6)

"""

def create_year_month():
    if not wd.db("REPORTSVR_VDPUTP01").empty:
        wd.open("REPORTSVR_VDPUTP01")
        wd.addCol("YEARMONTH", lambda row: row.DATE_CREATED.strftime("%Y%m%d")[0:6])

def extract_month():
    if not wd.db("REPORTSVR_VDPUTP01").empty:
        wd.open("REPORTSVR_VDPUTP01")
        wd.extract("Created_This_Month", filter=f"UTCID != '' and YEARMONTH == '{v_YearMonth}'")

def join_kyc():
    if not wd.db("Created_This_Month").empty:
        wd.open("Created_This_Month")
        wd.extract("Created_This_Month_Join", cols=["DATE_CREATED","BRANCHCD1","HOLDRNAME","OCCUP"	,"HOLDRTYP"	,"BIRTHDAT"	,"ADDRESS1","ADDRESS2","ADDRESS3","UTCID"])
        
        wd.open("Created_This_Month_Join")
        wd.join("KYC_Status_New_Customers", right=wd.db("Registrations"), how="left", on=["UTCID"])

def join_branch():
    if not wd.db("KYC_Status_New_Customers").empty:
        wd.open("KYC_Status_New_Customers")
        wd.join("MR2_INT", right=wd.db("REPORTSVR.VDPBRANCH")[["BRNCHCODE", "LONGDES"]], how="left", left_on=["BRANCHCD1"], right_on=["BRNCHCODE"])

def join_portal():
    if not wd.db("MR2_INT").empty:
        wd.open("MR2_INT")
        wd.join("MR2", right=wd.db("RISK RATINGS")[["NEW_UTCID", "RATING_SOURCE", "RISK_RATING"]], how="left", left_on=["UTCID"], right_on=["NEW_UTCID"])

def rename_fields():
    if not wd.db("MR2").empty:
        wd.open("MR2")
        wd.renameCol(columns={
            "LONGDES": "BRANCH_NAME",
            "BIRTHDAT": "DATE_OF_BIRTH",
            "BRANCHCD1": "CUSTOMER_BRANCH",
            "HOLDRNAME": "HOLDER_NAME",
        })
        wd.addCol("HOLDER_TYPE", lambda row: 'CORPORATION' if row.HOLDRTYP == 'C' else ('INDIVIDUAL' if row.HOLDRTYP == 'I' else ''))
        wd.addCol("OCCUPATION", lambda _: '')
        wd.addCol("POWEROFATTORNEYEXPIRE_DATE", lambda row: row.POWEROFATTORNEYEXPIRE.strftime("%Y%m%d"))

def export_details():
    if not wd.db("MR2").empty:
        wd.open("MR2")
        wd.extract("MR2_Export", cols=[ "UTCID", "DATE_CREATED", "HOLDER_NAME", "HOLDER_TYPE"	, "MISSING_KYC_IND", "BRANCH_NAME", "RATING_SOURCE", "RISK_RATING", "OCCUPATION", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DATE_OF_BIRTH", "ANNUALINCOME", "COUNTRYMULTIPLECITIZENSHIP", "COUNTRYOFBIRTH", "COUNTRYOFCITIZENSHIP", "COUNTRYOFRESIDENCE", "CREATED_DATE", "DOB_DATE", "DPNO", "DPNOCOUNTRY", "DPNOEXPIRE_DATE", "EMPLOYERNAME", "EMPLOYMENTSTATUS", "HOMEOWNERSHIP", "IDNO", "IDNOCOUNTRY", "IDNOEXPIRE_DATE", "MARITALSTATUS", "MULTIPLECITIZENSHIP", "PHONECONTACT", "PHONEHOME", "PHONEMOBILE", "PLACEOFWORK", "POWEROFATTORNEY", "POWEROFATTORNEYEXPIRE_DATE", "PPNO", "PPNOCOUNTRY", "PPNOEXPIRE_DATE", "PURPOSEOFRELATIONSHIP", "SEX", "TAXPAYERIDNO", "TAXPAYERIDNOCOUNTRY", "WORKCONTACTNO"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'MR2.mdb'))

def cleanup():
    if wd.tblName:
        wd.close()
    wd.delete("Created_This_Month")
    wd.delete("KYC_Status_New_Customers")
    wd.delete("MR2")
    wd.delete("MR2_INT")
    wd.delete("MR2_INT2")


if __name__ == "__main__":
    wd = DataAnalytics()

    create_year_month()
    extract_month()
    join_kyc()
    join_branch()
    join_portal()
    rename_fields()
    export_details()
    cleanup()



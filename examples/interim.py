import datetime as dt
from os.path import join, abspath
from DataAnalytics import DataAnalytics
PARENT_DIR = abspath()

e_Risk_Rating_Sheet_Name = ""
e_US_Rate = 0

# Append Field
def AddFields():
	if not wd.db("REPORTSVR.VDPUTR04N").empty:
		wd.open("REPORTSVR.VDPUTR04N")
		wd.addCol("NARR", lambda row: "")
	if not wd.db("REPORTSVR.VDPUTR081").empty:
		wd.open("REPORTSVR.VDPUTR081")
		wd.addCol("NARR", lambda row: "")
	if not wd.db("REPORTSVR.VDPUTR021IMD").empty:
		wd.open("REPORTSVR.VDPUTR021")
		wd.addCol("NARR", lambda row: "")
	if not wd.db("REPORTSVR.VDPUTR011").empty:
		wd.open("REPORTSVR.VDPUTR011")
		wd.addCol("NARR", lambda row: "")
	if not wd.db("REPORTSVR.VDPUCID01").empty:
		wd.open("REPORTSVR.VDPUCID01")
		wd.renameCol(columns={"BIRTHDAT_DATE": "BIRTHDAT"})


def AExt_Master():
	if not wd.db("REPORTSVR.VDPUTP01").empty:
		wd.open("REPORTSVR.VDPUTP01")
		wd.extract("Master_Funds", cols=["UTCID", "TRUSTCOD", "ACCTNOP"])


def BACreateDate():
	if not wd.db("REPORTSVR.VDPUCID01").empty:
		wd.open("REPORTSVR.VDPUCID01")
		wd.addCol("DATE_CREATED", lambda row: f"{dt.datetime.strptime(row.RECDATE1_DATE.strip()[:8], "%Y%m%d")}")


def BAJoinCustomers():
	if not wd.db("dbo.Registrations").empty:
		wd.open("dbo.Registrations")
		wd.join("Registrations", right=wd.db("REPORTSVR.VDPUCID01")[["UTCID", "ADDRESS1", "ADDRESS2", "ADDRESS3", "BIRTHDAT", "HOLDRTYP", "FIRSTNAME", "SURNAME"]], how="left", on=["UTCID"])


def BACreateKYCFields():
	if not wd.db("Registrations").empty:
		wd.open("Registrations")
		#Field created on March 21, 2015 to remove data errors. 		
		wd.addCol("PLACEOFWORK_NEW", lambda row: "" if row.PLACEOFWORK.upper() in ("N.A.", "N/A", "N\A", "Private", "Not APPLICABLE", "NA" , "NONE", "UNEMPLOYED", "UNKNOWN") else row.PLACEOFWORK)
		#Field created on March 21, 2015 to remove data errors. 	
		wd.addCol("EMPLOYERNAME_NEW", lambda row: "" if row.EMPLOYERNAME.upper() in ("N.A.", "N/A", "N\A", "Private", "Not APPLICABLE", "NA" , "NONE", "UNEMPLOYED", "UNKNOWN") else row.EMPLOYERNAME)

#Computes which KYC information is missing from records
#Commented equations reflect changes based on onsite UAT review in April 2015.
def BBCreateKYCFields():
	if not wd.db("Registrations").empty:
		wd.open("Registrations")
		wd.addCol("DATE_OF_BIRTH", lambda _: "BIRTHDAT")
		wd.addCol("CHECK_FNAME", lambda row: "First Name" if row.FIRSTNAME.strip() == "" and row.HOLDRTYP == "I" else "")
		wd.addCol("CHECK_SNAME", lambda row: "Last Name" if row.SURNAME.strip() == "" and row.HOLDRTYP == "I" else ("Business Name" if row.SURNAME.strip() == "" else ""))
		wd.addCol("CHECK_DOB", lambda row: "Date of Birth" if row.DATE_OF_BIRTH > dt.date.today() and row.HOLDRTYP != "C" else ("Incorporation Date" if row.DATE_OF_BIRTH > dt.date.today() else ""))
		wd.addCol("CHECK_SEX", lambda row: "Gender" if row.SEX.strip() == "" else "")
		wd.addCol("CHECK_ADDRESS1", lambda row: "Street Address" if row.ADDRESS1.strip() == "" else "")
		wd.addCol("CHECK_ADDRESS2", lambda row: "Address City" if row.ADDRESS2.strip() == "" else "")
		
		# Field updated March 21,2015 to exclude Check Address 3 if Check_COR is not blank.		
		wd.addCol("CHECK_ADDRESS3", lambda row: "Address Country" if row.COUNTRYOFRESIDENCE.strip() == "" and row.ADDRESS3.strip() == "" else "")
		wd.addCol("CHECK_ID", lambda row: "ID" if row.IDNO.strip() == "" and row.PPNO.strip() == "" and row.DPNO.strip() == "" and row.HOLDRTYP == "I" else ("ID" if row.PPNO.strip() == "" and row.DPNO.strip() == "" else ""))
		
		wd.addCol("CHECK_IDEXPIRE", lambda row: "")
		wd.addCol("CHECK_HOMEOWNERSHIP", lambda row: "")
		wd.addCol("CHECK_INCOME", lambda row: "")
		wd.addCol("CHECK_MULTIPLECITIZENSHIP", lambda row: "")
		wd.addCol("CHECK_COB", lambda row: "")
		wd.addCol("CHECK_CITIZENSHIP", lambda row: "")
		wd.addCol("CHECK_EMPLOYMENT", lambda row: "")
		wd.addCol("CHECK_OCCUPATION", lambda row: "")
		wd.addCol("CHECK_PHONE", lambda row: "")
		wd.addCol("CHECK_RELATIONSHIP", lambda row: "")
		wd.addCol("CHECK_POC", lambda row: "")
		wd.addCol("CHECK_EMPLOYER", lambda row: "")
		wd.addCol("MISSING_KYC_IND", lambda row: "")


#Create Fields
def CCreate_Fields():
	if not wd.db("REPORTSVR.VDPUCID01").empty:
		wd.open("REPORTSVR.VDPUCID01")
		wd.addCol("HOLDRNAME", lambda row: f"{row.FIRSTNAME.strip()} {row.SURNAME.strip()}")
		wd.addCol("DOB", lambda row: "BIRTHDAT")
	if not wd.db("REPORTSVR.VDPUTR04N").empty:
		wd.open("REPORTSVR.VDPUTR04N")
		wd.addCol("VALUNITS_New", lambda row: row.VALUNITS * e_US_Rate)
	if not wd.db("REPORTSVR.VDPURF031").empty:
		wd.open("REPORTSVR.VDPURF031")
		wd.addCol("TRUSTCOD", lambda _: "005")


#Rename fields for Append
def DRename_Fields():
	if not wd.db("REPORTSVR.VDPURF031").empty:
		wd.open("REPORTSVR.VDPURF031")
		wd.renameCol(columns={
			"TRNREF": "RECEIPTNO",
			"CUSNM": "CUSNAME"
		})
		wd.addCol("NARR", lambda row: "")

	if not wd.db("REPORTSVR.VDPUTR011").empty:
		wd.open("REPORTSVR.VDPUTR011")
		wd.addCol("AGTCOD", lambda row: "")
		wd.addCol("NARR", lambda row: "")
		wd.renameCol(columns={
			"TRNREF": "RECEIPTNO",
			"CUSNM": "CUSNAME",
			"PAYTYP": "PYMNTTYPE"
		})
	if not wd.db("REPORTSVR.VDPUTR021").empty:
		wd.open("REPORTSVR.VDPUTR021")
		wd.addCol("AGTCOD", lambda row: "")
		wd.addCol("NARR", lambda row: "")


def EExt_History():
	if not wd.db("REPORTSVR.VDPUTR04N").empty:
		wd.open("REPORTSVR.VDPUTR04N")
		wd.extract("DPUTR04N", cols=["TRUSTCOD", "ACCTNOP", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW"])
	if not wd.db("REPORTSVR.VDPURF031").empty:
		wd.open("REPORTSVR.VDPURF031")
		wd.extract("DPURF031", cols=["ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR"])
	if not wd.db("REPORTSVR.VDPUTR011").empty:
		wd.open("REPORTSVR.VDPUTR011")
		wd.extract("DPUTR011", cols=["ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR"])
	if not wd.db("REPORTSVR.VDPUTR021").empty:
		wd.open("REPORTSVR.VDPUTR021")
		wd.extract("DPUTR021", cols=["ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR"])
	if not wd.db("REPORTSVR.VDPUTR081").empty:
		wd.open("REPORTSVR.VDPUTR081")
		wd.extract("DPUTR081", cols=["ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR"])


#Append History Files
def FAppend_History():
	if not wd.db("DPUTR04N").empty:
		wd.open("DPUTR04N")
		wd.addCol("VALUNITS", lambda row: row.VALUNITS_NEW)

	# append dbs


# Join Transactions to Master to get UTCID
def GJoin_Hist_Master():
	if not wd.db("Master_History").empty:
		wd.open("Master_History")
		wd.join("Master_History_File", right=wd.db("Master_Funds")[["UTCID", "TRUSTCOD", "ACCTNO"]], how="left", on=["TRUSTCOD", "ACCTNO"])


# Join Transactions to Customers
def HJoinCustomers():
	if not wd.db("Master_History_File").empty and not wd.db("REPORTSVR.VDPUCID01").empty:
		wd.open("Master_History_File")
		wd.join("Transaction_History_INT", right=wd.db("REPORTSVR.VDPUCID01")[["UTCID", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DOB", "OCCUP", "HOLDRTYP", "HOLDRNAME", "BRANCHCD1"]], how="left", on=["UTCID"])

	# export dbs

def IJoinBranch():
	if not wd.db("Transaction_History_lite").empty and not wd.db("REPORTSVR.VDPBRANCH").empty:
		wd.open("Transaction_History_lite")
		wd.join("Transaction_History", right=wd.db("REPORTSVR.VDPBRANCH")[["BRNCHCODE", "LONGDES"]], how="left", left_on=["BRANCHCODE"], right_on=["BRNCHCODE"])


# Append Field
def JCreate_Fields():
	if not wd.db("Transaction_History").empty:
		wd.open("Transaction_History")
		wd.addCol("PAYMENT_TYPE", lambda row: "")
		wd.addCol("TRANSACTION_CHANNEL", lambda row: "")
		wd.addCol("TRANSACTION_TYPE", lambda row: "")
		wd.addCol("HOLDER_TYPE", lambda row: "")


# Rename Fields
def KRename_SourceFields():
	if not wd.db("Transaction_History").empty:
		wd.open("Transaction_History")
		wd.renameCol(columns={"LONGDES": "BRANCH_DESCRIPTION"})
		wd.renameCol(columns={"BRANCHCODE": "TRANSACTION_BRANCH"})
		wd.renameCol(columns={"BRANCHCD1": "CUSTOMER_BRANCH"})
		wd.renameCol(columns={"AGTCOD": "AGENT_CODE"})
		wd.renameCol(columns={"UNITPRC": "UNIT_PRICE"})
		wd.renameCol(columns={"NUMUNITS": "UNITS"})
		wd.renameCol(columns={"NARR": "NARRATIVE"})
		wd.renameCol(columns={"DOB": "DATE_OF_BIRTH"})
		wd.renameCol(columns={"HOLDRNAME": "HOLDER_NAME"})
		wd.renameCol(columns={"ACCTNO": "ACCT_NO"})
		wd.renameCol(columns={"CUSNAME": "CUSTOMER_NAME"})
		wd.renameCol(columns={"TRUSTCOD": "TRUST_CODE"})
		wd.renameCol(columns={"VALUNITS": "TRANSACTION_AMOUNT"})
		wd.addCol("OCCUPATION", lambda row: "")


#Create Date Fields
def LCreate_DateFields():
	if not wd.db("Transaction_History").empty:
		wd.open("Transaction_History")
		wd.addCol("POST_DATE", lambda row: "")
		wd.addCol("TRAN_DATE", lambda row: "")


#Amended April 17, 2015 to rename output file to make the changes below.
def MExt_Physical():
	if not wd.db("Transaction_History").empty:
		wd.open("Transaction_History")
		wd.extract("History_Transaction_Hist", cols=["UTCID", "TRUSTCOD", "ACCTNOP", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UNITPRC", "VALUNITS_NEW", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "ACCTNO", "TRUSTCOD", "TRNDAT", "POSTDAT", "TRNTYP", "RECEIPTNO", "UNITPRC", "NUMUNITS", "VALUNITS", "PYMNTTYPE", "CUSNAME", "AGTCOD", "BRANCHCODE", "TRANSRC", "TRANSTATUS", "NARR", "UTCID", "CUSTOMER_BRANCH", "ACCT_NO", "TRUST_CODE", "PAYMENT_TYPE", "TRANSACTION_CHANNEL", "TRANSACTION_BRANCH", "BRANCH_DESCRIPTION", "TRANSACTION_TYPE", "HOLDER_TYPE", "POST_DATE", "TRAN_DATE", "UNITS", "RECEIPTNO", "UNIT_PRICE", "TRANSACTION_AMOUNT", "CUSTOMER_NAME", "AGENT_CODE", "NARRATIVE", "HOLDER_NAME", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DATE_OF_BIRTH", "OCCUPATION", "TRANSRC", "TRANSTATUS"])


#Created April 17 to create Risk Ratings for all customers to include in all Results
def MFJoinPortaltoRisk():
	if not wd.db("dbo.EvaluationMatrix").empty:
		wd.open("dbo.EvaluationMatrix")
		wd.join("Risk_Ratings_and_PORTAL", right=wd.db("Risk_Ratings-")[["ADDRESS1", "ADDRESS2", "ADDRESS3", "BIRTHDAT", "HOLDRTYP", "FIRSTNAME", "SURNAME", "UTCID", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DOB", "OCCUP", "HOLDRTYP", "HOLDRNAME", "BRANCHCD1", "LONGDES", "UTCID", "NAME"]], how=WI_JOIN_ALL_REC, on=["UTCID", "UTCID", "A"])
		# import odbc: {'close_all': ''}


#Created April 17, 2015 to get Customer Branch Name
def MG_CreateKeys():
	if not wd.db("RISK_RATINGS").empty:
		wd.open("RISK_RATINGS")
		wd.addCol("RATING_SOURCE", lambda row: "")
		wd.addCol("RISK_RATING", lambda row: "")
		wd.addCol("NEW_UTCID", lambda row: "")


#Created April 17, 2015 to get Risk Rating and Source for each customer
def MHJoin_Portal():
	if not wd.db("History_Transaction_Hist").empty and not wd.db("RISK_RATINGS").empty:
		wd.open("History_Transaction_Hist")
		wd.join("History_Transaction_Hist_Risk", right=wd.db("RISK RATINGS")[["ADDRESS1", "ADDRESS2", "ADDRESS3", "BIRTHDAT", "HOLDRTYP", "FIRSTNAME", "SURNAME", "UTCID", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DOB", "OCCUP", "HOLDRTYP", "HOLDRNAME", "BRANCHCD1", "LONGDES", "UTCID", "NAME", "RATING_SOURCE", "RISK_RATING"]], how=WI_JOIN_ALL_IN_PRIM, on=["UTCID", "UTCID", "A"])


#Created April 17, 2015 to get Customer Branch Name
def MI_JoinCustomerBranch():
	if not wd.db("History_Transaction_Hist_Risk").empty:
		wd.open("History_Transaction_Hist_Risk")
		wd.join("History_Transaction_History", right=wd.db("REPORTSVR.VDPBRANCH")[["ADDRESS1", "ADDRESS2", "ADDRESS3", "BIRTHDAT", "HOLDRTYP", "FIRSTNAME", "SURNAME", "UTCID", "ADDRESS1", "ADDRESS2", "ADDRESS3", "DOB", "OCCUP", "HOLDRTYP", "HOLDRNAME", "BRANCHCD1", "LONGDES", "UTCID", "NAME", "RATING_SOURCE", "RISK_RATING", "LONGDES"]], how=WI_JOIN_ALL_IN_PRIM, on=["CUSTOMER_BRANCH", "BRNCHCODE", "A"])


#Amended April 17, 2015 to include Customer Branch Name
#Extract Daily Transactions
def NExtHist_Daily():
	if not wd.db("History_Transaction_History").empty:
		wd.open("History_Transaction_History")
		wd.renameCol(columns={"RECEIPTNO": "RECEIPT_NUMBER"})
		wd.renameCol(columns={"LONGDES": "BRANCH_NAME"})
	if not wd.db("History_Transaction_History").empty:
		wd.open("History_Transaction_History")
		wd.extract("Daily_Transactions_Today", cols=all)


#Extract History Transactions for Average
def OExtHist_Average():
	if not wd.db("History_Transaction_History").empty:
		wd.open("History_Transaction_History")
		wd.extract("Tran_Hist_Average", cols=all)


#Extract Daily New Accounts
def QCreateAccountDate():
	if not wd.db("REPORTSVR.VCRDHLDR").empty:
		wd.open("REPORTSVR.VCRDHLDR")
		wd.addCol("STATUS", lambda row: "")
		wd.addCol("CHGDATE", lambda row: "")
	if not wd.db("REPORTSVR.VDPUTP01").empty:
		wd.open("REPORTSVR.VDPUTP01")
		#	field.Equation = ""

		wd.addCol("CREATE_DATE", lambda row: "")
	if not wd.db("REPORTSVR.VPAYEE").empty:
		wd.open("REPORTSVR.VPAYEE")
		wd.renameCol(columns={"PAYEE_NAME": "PAYEENAME"})


def RExtAccounts_Daily():
	if not wd.db("REPORTSVR.VDPUTP01").empty:
		wd.open("REPORTSVR.VDPUTP01")
		wd.extract("Accounts_Created_Today", cols=all)


def SCleanup():
	if wd.tblName:
		wd.close()

	wd.delete("Master_Funds")
	wd.delete("Master_History")
	wd.delete("DPUTR04N")
	wd.delete("DPURF031")
	wd.delete("DPUTR011")
	wd.delete("DPUTR021")
	wd.delete("DPUTR081")
	wd.delete("Master_History_File")
	wd.delete("Transaction_History_INT")
	wd.delete("Transaction_History")
	wd.delete("History_Transaction_Hist")
	wd.delete(f"Risk_Ratings-{e_Risk_Rating_Sheet_Name.strip()}")
	wd.delete("Risk Ratings and PORTAL")
	wd.delete("History_Transaction_Hist_Risk")


if __name__ == "__main__":
	wd = DataAnalytics()

	AddFields()
	AExt_Master()
	BACreateDate()
	BAJoinCustomers()
	BACreateKYCFields()
	BBCreateKYCFields()
	CCreate_Fields()
	DRename_Fields()
	EExt_History()
	FAppend_History()
	GJoin_Hist_Master()
	HJoinCustomers()
	IJoinBranch()
	SCleanup()



import datetime as dt
from os.path import join, abspath
from DataAnalytics import DataAnalytics
PARENT_DIR = abspath()

e_Personal_Combine_Thresh = 0
e_Personal_combine_ratio = 0

# Get all Personal cash in and out for the fortnight.
def A_Get_CashInOut():
	wd.open("FORTNIGHTLY_TRANSACTIONS")
	wd.addCol("RUN_DATE", lambda _: dt.date.today())
	if not wd.db("FORTNIGHTLY_TRANSACTIONS").empty:
		wd.open("FORTNIGHTLY_TRANSACTIONS")
		wd.extract("Personal_Cash_In_Out", cols=["SOURCE_ACCOUNT", "DEBIT_CREDIT", "CURRENCY", "OTHER_PARTY_ACC", "EXCLUDE_FROM_PROFILING", "DATE", "TRANSACTION_CODE", "AMOUNT", "ORIGINAL_AMOUNT", "REFERENCE", "DESCRIPTION", "FROM_CUSTOMER", "TO_ACCOUNT", "ROLE_TYPE", "CUSTOMER_NUMBER", "NAME", "COUNTRY_RESIDENCE", "COUNTRY_NATIONALITY", "DATEOFBIRTH", "ADDRESS_LINE_1", "ADDRESS_LINE_2", "ADDRESS_LINE_3", "ADDRESS_LINE_4", "ADDRESS_LINE_5", "GENDER", "CUSTOMER_TYPE", "BUSINESS_CODE", "SSN", "COUNTRY_BIRTH", "OCCUPATION", "IDENTIFICATION_TYPE", "IDENTIFICATION_OTHER", "IDENTIFICATION_NUMBER", "ISSUING_AUTHORITY", "CONTACT_NUMBER", "COUNTRY_RISK", "CUSTOMER_STATUS", "RISK_CATEGORY", "PEP", "ACCOUNT_TYPE", "STATUS", "ACCOUNT_NAME", "TRANSACTION_CURRENCY", "TRANSACTION_AMOUNT", "ORIGINAL_TRANSACTION_AMOUNT", "RUN_DATE"], filter="TRANSACTION_CODE in ('CASH', 'ATMC', 'DBTC') and CUSTOMER_TYPE == 'P")

# Get the first cash in generated per account
def B_Get_First_CashIn_Per_Accont():
	pass

# Link cash In and out to place the first cash in date beside all the txns.
def C_Join_FirstCashIn_To_AllCashInOut():
	if not wd.db("Personal_Cash_In_Out").empty and not wd.db("First_Personal_Cash_In").empty:
		wd.open("Personal_Cash_In_Out")
		wd.join("Personal_Cash_InOut_First_CashIn", right=wd.db("First_Personal_Cash_In")[["SOURCE_ACCOUNT", "DATE"]], how="inner", on=["SOURCE_ACCOUNT"])

# Get first Personal cash out after first cash in.
def D_Get_First_CashOut_After_CashIn():
	pass

# Link Cash In and Out to cash out file
def E_Get_CashIn_CashOut_Details():
	if not wd.db("Personal_Cash_In_Out").empty and not wd.db("Personal_Cash_Out_After_In").empty:
		wd.open("Personal_Cash_In_Out")
		wd.join("Personal_Cash_In_Out_Temp2", right=wd.db("Personal_Cash_Out_After_In")[["SOURCE_ACCOUNT", "DATE1"]], how="inner", on=["SOURCE_ACCOUNT"])
	
	# Ensure ONLY records that fall after the first cash in are selected for the summary	
	if not wd.db("Personal_Cash_In_Out_Temp2").empty:
		wd.open("Personal_Cash_In_Out_Temp2")
		wd.extract("Personal_Cash_In_Out_Temp", filter="DATE >= DATE1")


def F_ModifyFieldDetails():
	if not wd.db("Personal_Cash_In_Out_Temp").empty:
		wd.open("Personal_Cash_In_Out_Temp")
		wd.renameCol(columns={
			"SOURCE_ACCOUNT": "ACCOUNT_NUMBER",
			"DEBIT_CREDIT": "DIRECTION_OF_TRANSACTION",
			"CURRENCY": "ORIGINAL_CURRENCY",
			"NAME": "CUSTOMER_NAME",
			"DATE": "TRANSACTION_DATE",
			"STATUS": "ACCOUNT_STATUS"
		})

def G_GetSummaryRecords():
	if not wd.db("Personal_Cash_In_Out_Temp").empty:
		wd.open("Personal_Cash_In_Out_Temp")
		wd.summby("Personal_Cash_In_Out_Summ", cols=["ACCOUNT_NUMBER"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "ACCOUNT_NUMBER": ["count"]})
		wd.renameCol(columns={"ACCOUNT_NUMBER_count": "NO_OF_RECS"})

	if not wd.db("Personal_Cash_In_Out_Summ").empty:
		wd.open("Personal_Cash_In_Out_Summ")
		wd.extract("Personal_Cash_In_Out_Summary", filter=f"{(e_Personal_Combine_Thresh + (e_Personal_combine_ratio * e_Personal_Combine_Thresh))} <= TRANSACTION_AMOUNT_sum >= {(e_Personal_Combine_Thresh - (e_Personal_combine_ratio * e_Personal_Combine_Thresh))}")

def H_GetDetailRecords():
	if not wd.db("Personal_Cash_In_Out_Temp").empty and not wd.db("Personal_Cash_In_Out_Summary").empty:
		wd.open("Personal_Cash_In_Out_Temp")
		wd.join("Personal_Cash_In_Out_Details", right=wd.db("Personal Cash In Out Summary")[["RUN_DATE", "ACCOUNT_NUMBER"]], how="inner", on=["ACCOUNT_NUMBER"])

def I_ModifyFieldSummary():
	if not wd.db("Personal_Cash_In_Out_Summary").empty:
		wd.open("Personal_Cash_In_Out_Summary")
		wd.renameCol(columns={"NO_OF_RECS": "NO_OF_TRANSACTIONS",
						"TRANSACTION_AMOUNT_sum": "TOTAL_TRANSACTION_AMOUNT"})

def J_Get_Total_Credit_Debit():
	if not wd.db("Personal_Cash_In_Out_Details").empty:
		wd.open("Personal_Cash_In_Out_Details")
		wd.summBy("Debit_Total_summ", ["ACCOUNT_NUMBER"], agg_funcs={key: ['sum'] if key != "ACCOUNT_NUMBER" else ['count'] for key in ["TRANSACTION_AMOUNT", "ACCOUNT_NUMBER"]})
		wd.renameCol(columns={"ACCOUNT_NUMBER_count": "NO_OF_RECS"})
		wd.extract("Debit_Total", filter="DIRECTION_OF_TRANSACTION == 'D'")
		
		wd.open("Personal_Cash_In_Out_Details")
		wd.summBy("Credit_Total_summ", ["ACCOUNT_NUMBER"], agg_funcs={key: ['sum'] if key != "ACCOUNT_NUMBER" else ['count'] for key in ["TRANSACTION_AMOUNT", "ACCOUNT_NUMBER"]})
		wd.renameCol(columns={"ACCOUNT_NUMBER_count": "NO_OF_RECS"})
		wd.extract("Credit_Total", filter="DIRECTION_OF_TRANSACTION == 'C'")

	if not wd.db("Credit_Total").empty:
		wd.open("Credit_Total")
		wd.renameCol(columns={"TRANSACTION_AMOUNT_sum": "TOTAL_CREDIT"})
	if not wd.db("Debit_Total").empty:
		wd.open("Debit_Total")
		wd.renameCol(columns={"TRANSACTION_AMOUNT_sum": "TOTAL_DEBIT"})

	if not wd.db("Personal_Cash_In_Out_Summary").empty and not wd.db("Credit_Total").empty and not wd.db("Debit_Total").empty:
		wd.open("Personal_Cash_In_Out_Summary")
		# visual connect: {'add_db': "Debit Total", 'add_assigns': [(id0, None), (id1, None), (id2, None)], 'master_db': id0, 'append_db_names': FALSE, 'include_all_prim_recs': FALSE, 'fields_to_include': [(id0, "ACCOUNT_NUMBER"), (id0, "NO_OF_TRANSACTIONS"), (id0, "TOTAL_TRANSACTION_AMOUNT"), (id0, "ACCOUNT_NAME"), (id0, "ACCOUNT_TYPE"), (id0, "ROLE_TYPE"), (id0, "ACCOUNT_STATUS"), (id0, "CUSTOMER_NUMBER"), (id0, "CUSTOMER_NAME"), (id0, "CUSTOMER_TYPE"), (id0, "TRANSACTION_DATE"), (id0, "RUN_DATE"), (id0, "TRANSACTION_CURRENCY"), (id1, "TOTAL_CREDIT"), (id2, "TOTAL_DEBIT")], 'create_virt_db': False, 'add_relation': [], 'db_name': "Personal Cash In Out Summary Final", 'output_db': dbName, 'perf_task': ''}


def K_Export():
	if not wd.db("Personal_Cash_In_Out_Summary_Final").empty:
		wd.open("Personal_Cash_In_Out_Summary_Final")
		wd.extract("Personal_Cash_In_Out_Summary_Final_export", cols=["ACCOUNT_NUMBER", "ACCOUNT_NAME", "ACCOUNT_TYPE", "ROLE_TYPE", "ACCOUNT_STATUS", "CUSTOMER_NUMBER", "CUSTOMER_NAME", "CUSTOMER_TYPE", "RUN_DATE", "NO_OF_TRANSACTIONS", "TOTAL_TRANSACTION_AMOUNT", "TRANSACTION_CURRENCY", "TOTAL_CREDIT", "TOTAL_DEBIT"])
		wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'Personal Cash In Out Summary.mdb'))

	if not wd.db("Personal_Cash_In_Out_Details").empty:
		wd.open("Personal_Cash_In_Out_Details")
		wd.extract("Personal_Cash_In_Out_Details_export", cols=["ACCOUNT_NUMBER", "ACCOUNT_NAME", "ACCOUNT_TYPE", "ROLE_TYPE", "ACCOUNT_STATUS", "CUSTOMER_NUMBER", "CUSTOMER_NAME", "CUSTOMER_TYPE", "RUN_DATE", "NO_OF_TRANSACTIONS", "TOTAL_TRANSACTION_AMOUNT", "TRANSACTION_CURRENCY", "TOTAL_CREDIT", "TOTAL_DEBIT", "ACCOUNT_NUMBER", "CUSTOMER_NUMBER", "TRANSACTION_DATE", "TRANSACTION_AMOUNT", "ORIGINAL_TRANSACTION_AMOUNT", "ORIGINAL_CURRENCY", "DIRECTION_OF_TRANSACTION", "TRANSACTION_CODE", "DESCRIPTION", "REFERENCE", "RUN_DATE"])
		wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'Personal Cash In Out Details.mdb'))

def Z_CleanUp():
	if wd.tblName:
		wd.close()
	wd.delete("Personal_Cash_In_Out")
	wd.delete("First_Personal_Cash_In")
	wd.delete("Personal_Cash_InOut_First_CashIn")
	wd.delete("Personal_Cash_Out_After_In")
	wd.delete("Personal_Cash_In_Out Details")
	wd.delete("Personal_Cash_In_Out Summary")
	wd.delete("Personal_Cash_In_Out Summ")
	wd.delete("Personal_Cash_In_Out Temp")
	wd.delete("Personal_Cash_In_Out_Temp2")
	wd.delete("Personal_Cash_In_Out Summary Final")
	wd.delete("Credit Total")
	wd.delete("Debit Total")


if __name__ == "__main__":
	wd = DataAnalytics()

	A_Get_CashInOut()
	B_Get_First_CashIn_Per_Accont()
	C_Join_FirstCashIn_To_AllCashInOut()
	D_Get_First_CashOut_After_CashIn()
	E_Get_CashIn_CashOut_Details()
	F_ModifyFieldDetails()
	G_GetSummaryRecords()
	H_GetDetailRecords()
	I_ModifyFieldSummary()
	J_Get_Total_Credit_Debit()
	K_Export()
	Z_CleanUp()
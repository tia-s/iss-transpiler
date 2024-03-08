import datetime as dt
from os.path import join, abspath
from DataAnalytics import DataAnalytics
PARENT_DIR = abspath()

   def A_Get_WireIn_ChequeOut():
# Get all Personal Cheque in and out for the fortnight.

wd.open("FORTNIGHTLY_TRANSACTIONS")
wd.addCol("RUN_DATE", lambda row: "")
if not wd.db("FORTNIGHTLY_TRANSACTIONS").empty:
wd.open("FORTNIGHTLY_TRANSACTIONS")
#		task.AddFieldToInc "CUS_NUMBER"

wd.extract("Personal_Wire_In_Cheque_Out", cols=['"SOURCE_ACCOUNT"', '"DEBIT_CREDIT"', '"CURRENCY"', '"OTHER_PARTY_ACC"', '"EXCLUDE_FROM_PROFILING"', '"DATE"', '"TRANSACTION_CODE"', '"AMOUNT"', '"ORIGINAL_AMOUNT"', '"REFERENCE"', '"DESCRIPTION"', '"FROM_CUSTOMER"', '"TO_ACCOUNT"', '"ROLE_TYPE"', '"CUSTOMER_NUMBER"', '"NAME"', '"COUNTRY_RESIDENCE"', '"COUNTRY_NATIONALITY"', '"DATEOFBIRTH"', '"ADDRESS_LINE_1"', '"ADDRESS_LINE_2"', '"ADDRESS_LINE_3"', '"ADDRESS_LINE_4"', '"ADDRESS_LINE_5"', '"GENDER"', '"CUSTOMER_TYPE"', '"BUSINESS_CODE"', '"SSN"', '"COUNTRY_BIRTH"', '"OCCUPATION"', '"IDENTIFICATION_TYPE"', '"IDENTIFICATION_OTHER"', '"IDENTIFICATION_NUMBER"', '"ISSUING_AUTHORITY"', '"CONTACT_NUMBER"', '"COUNTRY_RISK"', '"CUSTOMER_STATUS"', '"RISK_CATEGORY"', '"PEP"', '"ACCOUNT_TYPE"', '"STATUS"', '"ACCOUNT_NAME"', '"TRANSACTION_CURRENCY"', '"TRANSACTION_AMOUNT"', '"ORIGINAL_TRANSACTION_AMOUNT"', '"RUN_DATE"'])


# Get the first Wire in generated per account

def B_Get_First_WireIn_Per_Account():
if not wd.db("Personal_Wire_In_Cheque_Out").empty:
wd.open("Personal_Wire_In_Cheque_Out")
top recs: {'fields': ['"SOURCE_ACCOUNT"', '"DATE"'], 'keys': [('"SOURCE_ACCOUNT"', '"A"'), ('"DATE"', '"A"')], 'db_name': '"First Personal Wire In"', 'output_file': dbName, 'no_of_recs': 1, 'virt_db': False, 'perf_task': ''}


# Link Wire In and Cheque out to place the first Cheque in date beside all the txns.

def C_Join_FirstWireIn_To_AllWireInChequeOut():
if not wd.db("Personal_Wire_In_Cheque_Out").empty:
wd.open("Personal_Wire_In_Cheque_Out")
wd.join("Personal_Wire_In_Cheque_Out_First_WireIn", right=wd.db("First Personal Wire In")[['"DATE"']], how=WI_JOIN_MATCH_ONLY, on=['"SOURCE_ACCOUNT"', '"SOURCE_ACCOUNT"', '"A"'])


# Get first Personal Cheque out after first Wire in.

def D_Get_First_ChequeOut_After_CheckIn():
if not wd.db("Personal_Wire_In_Cheque_Out_First_WireIn").empty:
wd.open("Personal_Wire_In_Cheque_Out_First_WireIn")
top recs: {'fields': ['"SOURCE_ACCOUNT"', '"DATE"', '"SOURCE_ACCOUNT"', '"DATE"', '"DATE1"'], 'keys': [('"SOURCE_ACCOUNT"', '"A"'), ('"DATE"', '"A"'), ('"SOURCE_ACCOUNT"', '"A"'), ('"DATE"', '"A"')], 'db_name': '"Personal Cheque Out After Wire In"', 'output_file': dbName, 'no_of_recs': 1, 'virt_db': False, 'perf_task': ''}


# Link Cheque In and Out to Cheque out file

def E_Get_CheckIn_ChequeOut_Details():
if not wd.db("Personal_Wire_In_Cheque_Out").empty:
wd.open("Personal_Wire_In_Cheque_Out")
wd.join("Personal_Wire_In_Cheque_Out_Temp2", right=wd.db("Personal Cheque Out After Wire In")[['"DATE"', '"SOURCE_ACCOUNT"', '"DATE1"']], how=WI_JOIN_MATCH_ONLY, on=['"SOURCE_ACCOUNT"', '"SOURCE_ACCOUNT"', '"A"'])
# Ensure ONLY records that fall after the first Cheque in are selected for the summary	

if not wd.db("Personal_Wire_In_Cheque_Out_Temp2").empty:
wd.open("Personal_Wire_In_Cheque_Out_Temp2")
wd.extract("Personal_Wire_In_Cheque_Out_Temp", cols=all)


def F_ModifyFieldDetails():
if not wd.db("Personal_Wire_In_Cheque_Out_Temp").empty:
wd.open("Personal_Wire_In_Cheque_Out_Temp")
wd.renameCol(columns={"SOURCE_ACCOUNT": "ACCOUNT_NUMBER"})
wd.renameCol(columns={"DEBIT_CREDIT": "DIRECTION_OF_TRANSACTION"})
wd.renameCol(columns={"CURRENCY": "ORIGINAL_CURRENCY"})
wd.renameCol(columns={"NAME": "CUSTOMER_NAME"})
wd.renameCol(columns={"DATE": "TRANSACTION_DATE"})
wd.renameCol(columns={"STATUS": "ACCOUNT_STATUS"})


def G_GetSummaryRecords():
if not wd.db("Personal_Wire_In_Cheque_Out_Temp").empty:
wd.open("Personal_Wire_In_Cheque_Out_Temp")
wd.summBy(""Personal_Wire_In_Cheque_Out_Summ"", ['"ACCOUNT_NUMBER"'], agg_funcs={key: ['sum'] if key != ""ACCOUNT_NUMBER"" else ['count'] for key in ['"TRANSACTION_AMOUNT"', '"ACCOUNT_NUMBER"']})
wd.renameCol(columns={""ACCOUNT_NUMBER"_count": "NO_OF_RECS"})
wd.join(""Personal_Wire_In_Cheque_Out_Summ"", right=wd.db(""Personal_Wire_In_Cheque_Out_Summ"_summ")[['"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"TRANSACTION_DATE"', '"RUN_DATE"', '"TRANSACTION_CURRENCY"']], how="left")
if not wd.db("Personal_Wire_In_Cheque_Out_Summ").empty:
wd.open("Personal_Wire_In_Cheque_Out_Summ")
wd.extract("Personal_Wire_In_Cheque_Out_Summary", cols=all)


def H_GetDetailRecords():
if not wd.db("Personal_Wire_In_Cheque_Out_Temp").empty:
wd.open("Personal_Wire_In_Cheque_Out_Temp")
wd.join("Personal_Wire_In_Cheque_Out_Details", right=wd.db("Personal Cheque Out After Wire In")[['"DATE"', '"SOURCE_ACCOUNT"', '"DATE1"', '"RUN_DATE"']], how=WI_JOIN_MATCH_ONLY, on=['"ACCOUNT_NUMBER"', '"ACCOUNT_NUMBER"', '"A"'])


def I_ModifyFieldSummary():
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
wd.open("Personal_Wire_In_Cheque_Out_Summary")
wd.renameCol(columns={"NO_OF_RECS": "NO_OF_TRANSACTIONS"})
wd.renameCol(columns={"TRANSACTION_AMOUNT_SUM": "TOTAL_TRANSACTION_AMOUNT"})


def J_Get_Total_Credit_Debit():
if not wd.db("Personal_Wire_In_Cheque_Out_Details").empty:
wd.open("Personal_Wire_In_Cheque_Out_Details")
wd.summBy(""Debit_Total"", ['"ACCOUNT_NUMBER"'], agg_funcs={key: ['sum'] if key != ""ACCOUNT_NUMBER"" else ['count'] for key in ['"TRANSACTION_AMOUNT"', '"ACCOUNT_NUMBER"']})
wd.renameCol(columns={""ACCOUNT_NUMBER"_count": "NO_OF_RECS"})
if not wd.db("Personal_Wire_In_Cheque_Out_Details").empty:
wd.open("Personal_Wire_In_Cheque_Out_Details")
wd.summBy(""Credit_Total"", ['"ACCOUNT_NUMBER"'], agg_funcs={key: ['sum'] if key != ""ACCOUNT_NUMBER"" else ['count'] for key in ['"TRANSACTION_AMOUNT"', '"ACCOUNT_NUMBER"']})
wd.renameCol(columns={""ACCOUNT_NUMBER"_count": "NO_OF_RECS"})
if not wd.db("Credit_Total").empty:
wd.open("Credit_Total")
wd.renameCol(columns={"TRANSACTION_AMOUNT_SUM": "TOTAL_CREDIT"})
if not wd.db("Debit_Total").empty:
wd.open("Debit_Total")
wd.renameCol(columns={"TRANSACTION_AMOUNT_SUM": "TOTAL_DEBIT"})
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
if not wd.db(And).empty:
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
if not wd.db(And).empty:
if not wd.db("Credit_Total").empty:
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
if not wd.db(And).empty:
if not wd.db("Credit_Total").empty:
if not wd.db(And).empty:
if not wd.db("Personal_Wire_In_Cheque_Out_Summary").empty:
if not wd.db(And).empty:
if not wd.db("Credit_Total").empty:
if not wd.db(And).empty:
if not wd.db("Debit_Total").empty:
wd.open("Personal_Wire_In_Cheque_Out_Summary")
visual connect: {'add_db': '"Debit Total"', 'add_assigns': [(id0, None), (id1, None), (id2, None)], 'master_db': id0, 'append_db_names': FALSE, 'include_all_prim_recs': FALSE, 'fields_to_include': [(id0, '"ACCOUNT_NUMBER"'), (id0, '"NO_OF_TRANSACTIONS"'), (id0, '"TOTAL_TRANSACTION_AMOUNT"'), (id0, '"ACCOUNT_NAME"'), (id0, '"ACCOUNT_TYPE"'), (id0, '"ROLE_TYPE"'), (id0, '"ACCOUNT_STATUS"'), (id0, '"CUSTOMER_NUMBER"'), (id0, '"CUSTOMER_NAME"'), (id0, '"CUSTOMER_TYPE"'), (id0, '"TRANSACTION_DATE"'), (id0, '"RUN_DATE"'), (id0, '"TRANSACTION_CURRENCY"'), (id1, '"TOTAL_CREDIT"'), (id2, '"TOTAL_DEBIT"')], 'create_virt_db': False, 'add_relation': [], 'db_name': '"Personal Wire In Cheque Out Summary Final"', 'output_db': dbName, 'perf_task': ''}


def K_Export():
if not wd.db("Personal_Wire_In_Cheque_Out_Summary_Final").empty:
wd.open("Personal_Wire_In_Cheque_Out_Summary_Final")
export: {'fields': ['"ACCOUNT_NUMBER"', '"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"RUN_DATE"', '"NO_OF_TRANSACTIONS"', '"TOTAL_TRANSACTION_AMOUNT"', '"TRANSACTION_CURRENCY"', '"TOTAL_CREDIT"', '"TOTAL_DEBIT"'], 'perform_task': ''}
if not wd.db("Personal_Wire_In_Cheque_Out_Details").empty:
wd.open("Personal_Wire_In_Cheque_Out_Details")
export: {'fields': ['"ACCOUNT_NUMBER"', '"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"RUN_DATE"', '"NO_OF_TRANSACTIONS"', '"TOTAL_TRANSACTION_AMOUNT"', '"TRANSACTION_CURRENCY"', '"TOTAL_CREDIT"', '"TOTAL_DEBIT"', '"ACCOUNT_NUMBER"', '"CUSTOMER_NUMBER"', '"TRANSACTION_DATE"', '"TRANSACTION_AMOUNT"', '"ORIGINAL_TRANSACTION_AMOUNT"', '"ORIGINAL_CURRENCY"', '"DIRECTION_OF_TRANSACTION"', '"TRANSACTION_CODE"', '"DESCRIPTION"', '"REFERENCE"', '"RUN_DATE"'], 'perform_task': ''}


def Z_CleanUp():
if wd.tblName:
	wd.close()

wd.delete("Personal Wire In Cheque Out")
wd.delete("First Personal Wire In")
wd.delete("Personal Wire In Cheque Out First WireIn")
wd.delete("Personal Cheque Out After Wire In")
wd.delete("Personal Wire In Cheque Out Details")
wd.delete("Personal Wire In Cheque Out Summary")
wd.delete("Personal Wire In Cheque Out Summ")
wd.delete("Personal Wire In Cheque Out Temp")
wd.delete("Personal Wire In Cheque Out Temp2")
wd.delete("Personal Wire In Cheque Out Summary Final")
wd.delete("Credit Total")
wd.delete("Debit Total")



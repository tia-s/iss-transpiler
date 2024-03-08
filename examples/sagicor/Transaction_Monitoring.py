import datetime as dt
from os.path import join, abspath
from DataAnalytics import DataAnalytics
PARENT_DIR = abspath()

   def GetPersonalCashOut():
if not wd.db("T24_Customer_Transactions_Complete").empty:
wd.open("T24_Customer_Transactions_Complete")
wd.extract("Personal_Customer", cols=all)
# Cash Out

if not wd.db("Personal_Customer").empty:
wd.open("Personal_Customer")
wd.extract("Personal_Cash_Out", cols=all)


def ModifyFieldDetails():
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"SOURCE_ACCOUNT": "ACCOUNT_NUMBER"})
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"DEBIT_CREDIT": "DIRECTION_OF_TRANSACTION"})
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"CURRENCY": "ORIGINAL_CURRENCY"})
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"NAME": "CUSTOMER_NAME"})
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"DATE": "TRANSACTION_DATE"})
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.renameCol(columns={"STATUS": "ACCOUNT_STATUS"})


def GetSummaryRecords():
if not wd.db("Personal_Cash_Out").empty:
wd.open("Personal_Cash_Out")
wd.summBy(""Personal_Cash_Out_Summ"", ['"ACCOUNT_NUMBER"'], agg_funcs={key: ['sum'] if key != ""ACCOUNT_NUMBER"" else ['count'] for key in ['"TRANSACTION_AMOUNT"', '"ACCOUNT_NUMBER"']})
wd.renameCol(columns={""ACCOUNT_NUMBER"_count": "NO_OF_RECS"})
wd.join(""Personal_Cash_Out_Summ"", right=wd.db(""Personal_Cash_Out_Summ"_summ")[['"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"TRANSACTION_DATE"', '"TRANSACTION_CURRENCY"']], how="left")


def tempjusttogetextract():
wd.open("Personal_Cash_Out_Summ")
wd.extract("Personal_Cash_Out_Summary", cols=all)


def GetDetailRecords():
if not wd.db("Personal_Cash_Out").empty:
if not wd.db("Personal_Cash_Out").empty:
if not wd.db(And).empty:
if not wd.db("Personal_Cash_Out").empty:
if not wd.db(And).empty:
if not wd.db("Personal_Cash_Out_Summary").empty:
wd.open("Personal_Cash_Out")
wd.join("Personal_Cash_Out_Results", right=wd.db("Personal Cash Out Summary")[['"ACCOUNT_NUMBER"']], how=WI_JOIN_MATCH_ONLY, on=['"ACCOUNT_NUMBER"', '"ACCOUNT_NUMBER"', '"A"'])


def ModifyFieldSummary():
if not wd.db("Personal_Cash_Out_Summary").empty:
wd.open("Personal_Cash_Out_Summary")
wd.renameCol(columns={"NO_OF_RECS": "NO_OF_TRANSACTIONS"})


def temptogetaddcol():
wd.open("Personal_Cash_Out_Summary")
wd.renameCol(columns={"TRANSACTION_AMOUNT_SUM": "TOTAL_TRANSACTION_AMOUNT"})


def Export():
if not wd.db("Personal_Cash_Out_Summary").empty:
wd.open("Personal_Cash_Out_Summary")
export: {'fields': ['"ACCOUNT_NUMBER"', '"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"TRANSACTION_DATE"', '"NO_OF_TRANSACTIONS"', '"TOTAL_TRANSACTION_AMOUNT"', '"TRANSACTION_CURRENCY"'], 'perform_task': ''}
if not wd.db("Personal_Cash_Out_Results").empty:
wd.open("Personal_Cash_Out_Results")
export: {'fields': ['"ACCOUNT_NUMBER"', '"ACCOUNT_NAME"', '"ACCOUNT_TYPE"', '"ROLE_TYPE"', '"ACCOUNT_STATUS"', '"CUSTOMER_NUMBER"', '"CUSTOMER_NAME"', '"CUSTOMER_TYPE"', '"TRANSACTION_DATE"', '"NO_OF_TRANSACTIONS"', '"TOTAL_TRANSACTION_AMOUNT"', '"TRANSACTION_CURRENCY"', '"ACCOUNT_NUMBER"', '"CUSTOMER_NUMBER"', '"TRANSACTION_DATE"', '"TRANSACTION_AMOUNT"', '"ORIGINAL_TRANSACTION_AMOUNT"', '"ORIGINAL_CURRENCY"', '"DIRECTION_OF_TRANSACTION"', '"TRANSACTION_CODE"', '"DESCRIPTION"', '"REFERENCE"'], 'perform_task': ''}


def CleanUp():
if wd.tblName:
	wd.close()

wd.delete("Personal Customer")
wd.delete("Personal Cash Out")
wd.delete("Personal Cash Out Results")
wd.delete("Personal Cash Out Summary")



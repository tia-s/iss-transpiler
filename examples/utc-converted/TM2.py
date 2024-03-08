import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_TM2_Threshold = 0

# Summarize Daily Cash Transactions
# Updated Dec 2018 to exclude Transfers and Repurchase Cancellation
def BSummDailyTrans():
    if not wd.db("Daily_Transactions_Today").empty:
        wd.open("Daily_Transactions_Today")
        wd.summBy("Summ_DAILY_TRAN_TM2_summ", cols=["UTCID", "POST_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})

        wd.join("Summ_DAILY_TRAN_TM2_join", right=wd.db("Daily_Transactions_Today")[["UTCID", "HOLDER_TYPE", "HOLDER_NAME", "CUSTOMER_BRANCH", "OCCUPATION", "RATING_SOURCE", "RISK_RATING", "BRANCH_NAME"]], how="left", on=["UTCID"])
        wd.extract("Summ_DAILY_TRAN_TM2", filter="PAYMENT_TYPE in ('CHEQUE', 'CASH') and TRANSACTION_TYPE == 'SALE' and RECEIPT_NUMBER not in ('TR', 'RC')")

# Extract Daily Transactiona above Threshold
def CExtResult_Summ():
    if not wd.db("Summ_DAILY_TRAN_TM2").empty:
        wd.open("Summ_DAILY_TRAN_TM2")
        wd.extract("TM2_Summ", filter=f"TRANSACTION_AMOUNT_sum >= {e_TM2_Threshold}")

# Join to Extract Transactions above Threshold at Detail
def EJoinExtResult_Dtl():
    if not wd.db("Daily_Transactions_Today").empty and not wd.db("TM2_Summ").empty:
        wd.open("Daily_Transactions_Today")
        wd.join("TM2_DTLS_join", right=wd.db("TM2_Summ")[["UTCID", "NO_OF_RECS"]], how="inner", on=["UTCID"])
        wd.extract("TM2_DTLS", filter="PAYMENT_TYPE in ('CHEQUE', 'CASH') and TRANSACTION_TYPE == 'SALE'")

def HRename_ResultFields():
    if not wd.db("TM2_SUMM").empty:
        wd.open("TM2_SUMM")
        wd.renameCol(columns={
            "NO_OF_RECS": "NO_OF_TRANSACTIONS",
            "TRANSACTION_AMOUNT_sum": "TRANSACTION_SUMMARY"
        })

def IExportDatabase():
    if not wd.db("TM2_DTLS").empty:
        wd.open("TM2_DTLS")
        wd.extract("TM2_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","RECEIPT_NUMBER","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","TRAN_DATE","TRANSACTION_AMOUNT","UNITS","UNIT_PRICE","PAYMENT_TYPE","TRANSACTION_CHANNEL","POST_DATE","AGENT_CODE","CUSTOMER_NAME","ADDRESS1","ADDRESS2","ADDRESS3","DATE_OF_BIRTH","OCCUPATION","HOLDER_NAME","NARRATIVE","NO_OF_RECS","HOLDER_TYPE","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM2_DTLS.mdb'))

    if not wd.db("TM2_SUMM").empty:
        wd.open("TM2_SUMM")
        wd.extract("TM2_SUMM_export", cols=["UTCID","CUSTOMER_BRANCH","HOLDER_NAME","HOLDER_TYPE","POST_DATE","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM2_SUMM.mdb'))

def JCleanUp():
    if wd.tblName:
        wd.close()
    wd.delete("Summ_DAILY_TRAN_TM2")
    wd.delete("TM2_SUMM")
    wd.delete("TM2_DTLS")

if __name__ == "__main__":
    wd = DataAnalytics()

    BSummDailyTrans()
    CExtResult_Summ()
    EJoinExtResult_Dtl()
    HRename_ResultFields()
    IExportDatabase()
    JCleanUp()
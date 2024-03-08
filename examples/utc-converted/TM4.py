import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_TM4_Trans_Thres = 0

# Summarize Daily Repo Transactions
def ASummDailyTransRP():
    if not wd.db("Weekly_History").empty:
        wd.open("Weekly_History")
        wd.summBy("Summ_DAILY_TRAN_TM4_summ", cols=["UTCID", "RUN_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS"
        })

        wd.join("Summ_DAILY_TRAN_TM4_join", right=wd.db("Weekly_History")[["UTCID", "POST_DATE","HOLDER_TYPE","CUSTOMER_BRANCH","HOLDER_NAME","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"]], how="left", on=["UTCID"])
        wd.extract("Summ_DAILY_TRAN_TM4", filter="TRANSACTION_TYPE == 'REPO' and (TRANSACTION_CHANNEL == 'WIRE' or (PAYMENT_TYPE == 'E' and TRANSRC == 'IC')) and TRANSACTION_BRANCH == '100'")

# Extract Daily Transactions above Threshold
def BExtResult_Summ():
    if not wd.db("Summ_DAILY_TRAN_TM4").empty:
        wd.open("Summ_DAILY_TRAN_TM4")
        wd.extract("TM4_Summ", filter=f"TRANSACTION_AMOUNT_sum >= {e_TM4_Trans_Thres}")

# Join to Extract Transactions above Threshold at Detail
def CJoinExtResult_Dtl():
    if not wd.db("Weekly_History").empty and not wd.db("TM4_Summ").empty:
        wd.open("Weekly_History")
        wd.join("TM4_DTLS_join", right=wd.db("TM4_Summ")[["UTCID", "NO_OF_RECS"]], how="inner", on=["UTCID"])
        wd.extract("TM4_DTLS", filter="TRANSACTION_TYPE == 'REPO' and (TRANSACTION_CHANNEL == 'WIRE' or (PAYMENT_TYPE == 'E' and TRANSRC == 'IC')) and TRANSACTION_BRANCH == '100'")

def FRename_ResultFields():
    if not wd.db("TM4_SUMM").empty:
        wd.open("TM4_SUMM")
        wd.renameCol(columns={
            "NO_OF_RECS": "NO_OF_TRANSACTIONS",
            "TRANSACTION_AMOUNT_sum": "TRANSACTION_SUMMARY"
        })

def GExportDatabase():
    if not wd.db("TM4_DTLS").empty:
        wd.open("TM4_DTLS")
        wd.extract("TM4_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","RECEIPT_NUMBER","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","TRAN_DATE","TRANSACTION_AMOUNT","UNITS","UNIT_PRICE","PAYMENT_TYPE","TRANSACTION_CHANNEL","POST_DATE","RUN_DATE","AGENT_CODE","CUSTOMER_NAME","ADDRESS1","ADDRESS2","ADDRESS3","DATE_OF_BIRTH","OCCUPATION","HOLDER_NAME","NARRATIVE","NO_OF_RECS","HOLDER_TYPE","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM4_DTLS.mdb'))

    if not wd.db("TM4_Summ").empty:
        wd.open("TM4_Summ")
        wd.extract("TM4_Summ_export", cols=["UTCID","CUSTOMER_BRANCH","HOLDER_TYPE","RUN_DATE","POST_DATE","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY","HOLDER_NAME","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM4_Summ.mdb'))

def HCleanUp():
    if wd.tblName:
        wd.close()
    wd.delete("Summ_DAILY_TRAN_TM4")
    wd.delete("TM4_SUMM")
    wd.delete("TM4_DTLS")
    
if __name__ == "__main__":
    wd = DataAnalytics()

    ASummDailyTransRP()
    BExtResult_Summ()
    CJoinExtResult_Dtl()
    FRename_ResultFields()
    GExportDatabase()
    HCleanUp()
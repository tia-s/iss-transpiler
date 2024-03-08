import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_TM1_SA_Threshold_Pct = 0
e_TM1_Min_Value = 0

# Summarize History to get Customer Average by Transaction Type
def ASummHist_Average():
    if not wd.db("Tran_Hist_Average").empty:
        wd.open("Tran_Hist_Average")
        wd.summBy("Summ_Hist_Average_summ", cols=["UTCID", "TRANSACTION_TYPE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum", "mean"], "UTCID": ["count"]})
        wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})
        wd.extract("Summ_Hist_Average", filter="TRANSACTION_TYPE in ('SALE', 'REPO') and PAYMENT_TYPE != 'BALANCE' and TRANSACTION_AMOUNT != 0.00")

# Create Join Key in Today's Database
def CCreateSumm_Today():
    if not wd.db("Daily_Transactions_Today").empty:
        wd.open("Daily_Transactions_Today")
        wd.summBy("Summ_Tran_Today_summ", cols=["UTCID", "TRANSACTION_TYPE", "POST_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={"UTCID_count": "NO_OF_RECS1"})
    wd.join("Summ_Tran_Today_join", right=wd.db("Daily_Transactions_Today")[["UTCID","CUSTOMER_BRANCH","HOLDER_TYPE","HOLDER_NAME","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"]], on=["UTCID"], how="left")
    wd.extract("Summ_Tran_Today", filter="UTCID != '' and TRANSACTION_TYPE in ('SALE', 'REPO') and PAYMENT_TYPE != 'BALANCE' and TRANSACTION_AMOUNT != 0.00")
    if not wd.db("Summ_Tran_Today").empty:
        wd.renameCol(colums={"TRANSACTION_AMOUNT_sum": "TRANSACTION_SUMMARY"})

# Join to get File to apply criteria to get results.
def DExtResults_INT():
    if not wd.db("Summ_Tran_Today").empty and not wd.db("Summ_Hist_Average").empty:
        wd.open("Summ_Tran_Today")
        wd.join("TM1_INT", right=wd.db("Summ_Hist_Average")[["UTCID", "TRANSACTION_AMOUNT_AVERAGE", "TRANSACTION_AMOUNT_SUM"]], how="left", on=["UTCID"])

# Extract exceptions based on criteria.
def EExtResults():
    if not wd.db("TM1_INT").empty:
        wd.open("TM1_INT")
        wd.extract("TM1_SUMM", filter=f"TRANSACTION_TYPE == 'SALE' and TRANSACTION_SUMMARY > (TRANSACTION_AMOUNT_AVERAGE * {e_TM1_SA_Threshold_Pct}) and TRANSACTION_SUMMARY > {e_TM1_Min_Value} and TRANSACTION_AMOUNT_AVERAGE > {e_TM1_Min_Value}")

# Analysis: Summarization
def GSumm_Dtls():
    if not wd.db("Daily_Transactions_Today").empty and not wd.db("TM1_SUMM").empty:
        wd.open("Daily_Transactions_Today")
        wd.join("TM1_DTLS_join", right=wd.db("TM1_SUMM")[["UTCID", "TRANSACTION_AMOUNT_SUM", "TRANSACTION_AMOUNT_AVERAGE"]], how="inner", on=["UTCID"])
        wd.extract("TM1_DTLS", filter="TRANSACTION_TYE in ('SALE', 'REPO') and PAYMENT_TYPE != 'BALANCE' AND TRANSACTION_AMOUNT != 0.00")

# Data: Index Database
def HRename_ResultFields():
    if not wd.db("TM1_SUMM").empty:
        wd.open("TM1_SUMM")
        wd.renameCol(columns={
            "NO_OF_RECS": "NO_OF_TRANSACTIONS",
            "TRANSACTION_AMOUNT_AVERAGE": "CUSTOMER_AVERAGE",
            "TRANSACTION_AMOUNT_SUM": "CUSTOMER_SUMMARY",
            "TRANSACTION_AMOUNT_AVERAGE": "CUSTOMER_AVERAGE"
        })

def IExportDatabase():
    if not wd.db("TM1_SUMM").empty:
        wd.open("TM1_SUMM")
        wd.extract("TM1_SUMM_export", cols=["UTCID","CUSTOMER_BRANCH","HOLDER_TYPE","POST_DATE","HOLDER_NAME","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY","TRANSACTION_TYPE","CUSTOMER_AVERAGE","CUSTOMER_SUMMARY","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM1_SUMM.mdb'))

    if not wd.db("TM1_DTLS").empty:
        wd.open("TM1_DTLS")
        wd.extract("TM1_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","PAYMENT_TYPE","TRANSACTION_CHANNEL","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","HOLDER_TYPE","POST_DATE","TRAN_DATE","UNITS","RECEIPT_NUMBER","UNIT_PRICE","TRANSACTION_AMOUNT","CUSTOMER_NAME","AGENT_CODE","NARRATIVE","HOLDER_NAME","ADDRESS1","ADDRESS2","ADDRESS3","DATE_OF_BIRTH","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM1_DTLS.mdb'))

def JCleanUp():
    if wd.tblName:
        wd.close()
    wd.delete("Summ_Hist_Average")
    wd.delete("Summ_Tran_Today")
    wd.delete("TM1_INT")
    wd.delete("TM1_SUMM")
    wd.delete("TM1_DTLS")

if __name__ == "__main__":
    wd = DataAnalytics()

    ASummHist_Average()
    CCreateSumm_Today()
    DExtResults_INT()
    EExtResults()
    GSumm_Dtls()
    HRename_ResultFields()
    IExportDatabase()
    JCleanUp()
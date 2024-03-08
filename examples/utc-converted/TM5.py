import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_TM5_Thresh = 0

# Extract Incoming Wires Exception
def ASummDailyTrans():
    if not wd.db("Daily_Transactions_Today").empty:
        wd.open("Daily_Transactions_Today")
        wd.summBy("TM5_Summ_INT_summ", cols=["UTCID", "POST_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS"
        })

        wd.join("TM5_Summ_INT_join", right=wd.db("Daily_Transactions_Today")[["UTCID", "HOLDER_TYPE"]], on=["UTCID"], how="left")
        wd.extract("TM5_Summ_INT", filter="TRANSACTION_TYPE == 'SALE' and (TRANSACTION_CHANNEL == 'WIRE' or (PAYMENT_TYPE == 'E' and TRANSRC == 'IC'))")

def BExtResult_Summ():
    if not wd.db("TM5_Summ_INT").empty:
        wd.open("TM5_Summ_INT")
        wd.extract("TM5_SUMM_INT2", filter=f"TRANSACTION_AMOUNT_sum >= {e_TM5_Thresh}")

# Join to Extract Transactions above Threshold at Detail
def DJoinExtResult_Dtl():
    if not wd.db("Daily_Transactions_Today").empty and not wd.db("TM5_Summ_INT2").empty:
        wd.open("Daily_Transactions_Today")
        wd.join("TM5_DTLS_join", right=wd.db("TM5_Summ_INT2")[["UTCID", "NO_OF_RECS"]], how="inner", on=["UTCID"])
        wd.extract("TM5_DTLS", filter="TRANSACTION_TYPE == 'SALE' and (TRANSACTION_CHANNEL == 'WIRE' or (PAYMENT_TYPE == 'E' and TRANSRC == 'IC'))")

def EJoinExtResult_Summ():
    if not wd.db("TM5_DTLS").empty:
        wd.open("TM5_DTLS")
        wd.summBy("TM5_Summ1_summ", cols=["UTCID", "POST_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS1"
        })
        wd.join("TM5_Summ1", right=wd.db("TM5_DTLS")[["UTCID","HOLDER_TYPE","HOLDER_NAME","CUSTOMER_BRANCH","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"]], on=["UTCID"], how="left")

def FRename_ResultFields():
    if not wd.db("TM5_SUMM1").empty:
        wd.open("TM5_SUMM1")
        wd.renameCol(columns={
            "NO_OF_RECS1": "NO_OF_TRANSACTIONS",
            "TRANSACTION_AMOUNT_sum": "TRANSACTION_SUMMARY"
        })

def GExportDatabase():
    if not wd.db("TM5_DTLS").empty:
        wd.open("TM5_DTLS")
        wd.extract("TM5_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","RECEIPT_NUMBER","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","TRAN_DATE","TRANSACTION_AMOUNT","UNITS","UNIT_PRICE","PAYMENT_TYPE","TRANSACTION_CHANNEL","HOLDER_TYPE","CUSTOMER_BRANCH","POST_DATE","AGENT_CODE","CUSTOMER_NAME","ADDRESS1","ADDRESS2","ADDRESS3","DATE_OF_BIRTH","OCCUPATION","HOLDER_NAME","NARRATIVE","NO_OF_RECS","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM5_DTLS.mdb'))

    if not wd.db("TM5_Summ").empty:
        wd.open("TM5_Summ")
        wd.extract("TM5_Summ_export", cols=["UTCID","CUSTOMER_BRANCH","HOLDER_TYPE","POST_DATE","HOLDER_NAME","OCCUPATION","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY","RATING_SOURCE","RISK_RATING","BRANCH_NAME","USERID","MONITORHIDDEN_WORKFLOWASSIGNEE"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM5_SUMM.mdb'))

def HCleanUp():
    if wd.tblName:
        wd.close()
    wd.delete("TM5_Summ_INT")
    wd.delete("TM5_SUMM_INT2")
    wd.delete("TM5_SUMM")
    wd.delete("TM5_SUMM1")
    wd.delete("TM5_DTLS")

if __name__ == "__main__":
    wd = DataAnalytics()

    ASummDailyTrans()
    BExtResult_Summ()
    DJoinExtResult_Dtl()
    EJoinExtResult_Summ()
    FRename_ResultFields()
    # Assign random items to users	
	# Client.RunIDEAScriptEx "AUTO-bp.ISS", "TM5_Summ1.IMD", "TM5_Summ.IMD", "", ""
    GExportDatabase()
    HCleanUp()
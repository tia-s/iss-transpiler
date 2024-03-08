import datetime as dt
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_Internal_List_Sheet_Name = ''
e_WL2_High_Thresh = 0
e_WL2_Medium_Thresh = 0

internal_watchlist_db_name = f"AML Compliance Internal Monitoring List-{e_Internal_List_Sheet_Name.strip()}"

# Join Today's Transactions to Internal Watch List
def ABJoinTransWatchlist():
    if not wd.db("Daily_Transactions_Today").empty and not wd.db(internal_watchlist_db_name).empty:
        wd.open("Daily_Transactions_Today")
        wd.join("WL2_DTLS_INT", right=wd.db(internal_watchlist_db_name), left_on=["UTCID"], right_on=["UTC_ID"], how="inner")

def ACRenameFields():
    if not wd.db("WL2_DTLS_INT").empty:
        wd.open("WL2_DTLS_INT")
        wd.renameCol(columns={
            "RISK_RATING1": "WL_RISK_RATING",
            "RISK_COMMENT": "WL_RISK_COMMENT"
        })

# Summarize WL2 Details
def BSummDetails():
    if not wd.db("WL2_DTLS_INT").empty:
        wd.open("WL2_DTLS_INT")
        wd.summBy("WL2_SUMM_INT_summ", cols=["UTCID", "POST_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS"
        })

        wd.join("WL2_SUMM_INT", right=wd.db("WL2_DTLS_INT")[["CUSTOMER_BRANCH","HOLDER_TYPE","CATEGORY","WL_RISK_RATING","WL_RISK_COMMENT","NAME_OF_CUSTOMER","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"]], how="left", on=["UTCID"])

def FRename_ResultFields():
    if not wd.db("WL2_SUMM_INT").empty:
        wd.open("WL2_SUMM_INT")
        wd.renameCol(columns={
            "NO_OF_RECS": "NO_OF_TRANSACTIONS",
            "TRANSACTION_AMOUNT_sum": "TRANSACTION_SUMMARY"
        })

# April 20 - Include new thresholds for Medium and High
def FBExtResult_Summ():
    # NB: the IDEA file had "WL2_Summ_INT" but that file wouldn't exist.
    if not wd.db("WL2_SUMM_INT").empty:
        wd.open("WL2_SUMM_INT")
        wd.extract("WL2_SUMM", filter=f"(WL_RISK_RATING == 'High' and TRANSACTION_SUMMARY == {e_WL2_High_Thresh}) or WL_RISK_RATING == 'Medium' and TRANSACTION_SUMMARY == {e_WL2_Medium_Thresh}")

def FCJoinExtResult_Dtl():
    if not wd.db("Daily_Transactions_Today").empty and not wd.db("WL2_SUMM").empty:
        wd.open("Daily_Transactions_Today")
        wd.join("WL2_DTLS", right=wd.db("WL2_SUMM")[["UTCID", "NO_OF_TRANSACTIONS"]], on=["UTCID"], how="inner")

def GExportDatabase():
    if not wd.db("WL2_DTLS").empty:
        wd.open("WL2_DTLS")
        wd.extract("WL2_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","RECEIPT_NUMBER","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","TRAN_DATE","TRANSACTION_AMOUNT","UNITS","UNIT_PRICE","PAYMENT_TYPE","TRANSACTION_CHANNEL","POST_DATE","AGENT_CODE","CUSTOMER_NAME","ADDRESS1","ADDRESS2","ADDRESS3","DATE_OF_BIRTH","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME","HOLDER_NAME","NARRATIVE","HOLDER_TYPE"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'WL2_DTLS.mdb'))

    if not wd.db("WL2_SUMM").empty:
        wd.open("WL2_SUMM")
        wd.extract("WL2_SUMM_export", cols=["UTCID","CUSTOMER_BRANCH","HOLDER_TYPE","POST_DATE","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY","CATEGORY","WL_RISK_RATING","WL_RISK_COMMENT","NAME_OF_CUSTOMER","OCCUPATION","RATING_SOURCE","RISK_RATING","BRANCH_NAME"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'WL2_SUMM.mdb'))

def JCleanUp():
    if wd.tblName:
        wd.close()
    wd.delete("WL2_SUMM")
    wd.delete("WL2_DTLS")
    wd.delete("WL2_SUMM_INT")
    wd.delete("WL2_DTLS_INT")

if __name__ == "__main__":
    wd = DataAnalytics()

    ABJoinTransWatchlist()
    ACRenameFields()
    BSummDetails()
    FRename_ResultFields()
    FBExtResult_Summ()
    FCJoinExtResult_Dtl()
    GExportDatabase()
    JCleanUp()
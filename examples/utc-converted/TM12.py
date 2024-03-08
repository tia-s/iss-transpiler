import datetime as dt
from dateutil.relativedelta import relativedelta
from os.path import join, abspath
import pypyodbc

from DataAnalytics import DataAnalytics

PARENT_DIR = abspath('')
e_TM12_Thresh = 0

def A_JoinCustTransactions():
    if not wd.db("History_Transaction_History").empty:
        wd.open("History_Transaction_History")
        wd.addCol("DIFFERENCE_BETWEEN_DATES", lambda row: relativedelta(dt.date.today(), row.TRAN_DATE).years)
        wd.extract("Monthly_History12", filter=f"UTCID != '' and TRANSACTION_TYPE in ('SALE', 'REPO') and PAYMENT_TYPE != 'BALANCE' and DIFFERENCE_BETWEEN_DATES < 31")

    if not wd.db("Monthly_History12").empty:
        wd.open("Monthly_History12")
        # NB: this part was commented out::
        # wd.addCol("FIRST_TWO_DIGITS_RECEIPT", lambda row: row.RECEIPT_NUMBER[:2])
        # wd.extract("Monthly_History_TM12", filter="FIRST_TWO_DIGITS_RECEIPT not in ('VR','VX','VG','VC','VS','VH','VT','VV','VB','VM','VN','VA','VP','VD','VJ','VM')")
        ######
        wd.extract("Monthly_History_TM12")
        wd.open("Monthly_History_TM12")
        wd.addCol("RUN_DATE", lambda _: dt.date.today())

    if not wd.db("Monthly_History_TM12").empty and not wd.db("REPORTSVR.VDPUCID01").empty:
        wd.open("Monthly_History_TM12")
        wd.join("Deposits_join", right=wd.db("REPORTSVR.VDPUCID01"), how="inner", on=["UTCID"])
        wd.extract("Deposits", filter="UTCID != '' and TRANSACTION_TYPE == 'SALE' and PAYMENT_TYPE != 'BALANCE' and CUSTOMER_NAME != 'UTC ATM'")

def B_SummWeeklyTransDay():
    if not wd.db("Deposits").empty:
        wd.open("Deposits")
        wd.summBy("TM12_SUMM_INT_sum", cols=["UTCID", "TRAN_DATE", "RUN_DATE"], agg_funcs={"TRANSACTION_AMOUNT": ["sum"], "UTCID": ["count"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS"
        })

        wd.join("TM12_SUMM_INT", right=wd.db("Deposits")[["UTCID","CUSTOMER_BRANCH","BRANCH_NAME","RATING_SOURCE","RISK_RATING","HOLDER_TYPE","HOLDER_NAME","OCCUPATION","STAFFIND"]], how="left", on=["UTCID"])

def C_SummWeeklyTransExclude():
    if not wd.db("TM12_SUMM_INT").empty:
        wd.open("TM12_SUMM_INT")
        wd.summBy("TM12_SUMM_INT2_summ", cols=["UTCID", "RUN_DATE"], agg_funcs={key: ["sum"] if key != "UTCID" else ["count"] for key in ["TRANSACTION_AMOUNT_sum", "NO_OF_RECS", "UTCID"]})
        wd.renameCol(columns={
            "UTCID_count": "NO_OF_RECS1"
        })

        wd.join("TM12_SUMM_INT2_join", right=wd.db("TM12_SUMM_INT")[["UTCID","CUSTOMER_BRANCH","BRANCH_NAME","RATING_SOURCE","RISK_RATING","HOLDER_TYPE","HOLDER_NAME","OCCUPATION","STAFFIND"]], how="left", on=["UTCID"])
        wd.extract("TM12_SUMM_INT2", filter=f"TRANSACTION_AMOUNT_sum <= {e_TM12_Thresh}")

def F_AboveDollarThreshold():
    if not wd.db("TM12_SUMM_INT2").empty:
        wd.open("TM12_SUMM_INT2")
        wd.extract("TM12_SUMM", filter=f"TRANSACTION_AMOUNT_sum_sum > {e_TM12_Thresh}")

def G_Rename_SummFields():
    if not wd.db("TM12_SUMM").empty:
        wd.open("TM12_SUMM")
        wd.renameCol(columns={
            "NO_OF_RECS_sum": "NO_OF_TRANSACTIONS",
            "NO_OF_RECS1": "TRANSACTION_DAYS",
            "TRANSACTION_AMOUNT_sum_sum": "TRANSACTION_SUMMARY",
        })
        wd.addCol("STAFF", lambda row: 'Yes' if row.STAFFIND.strip() == 'Y' else 'No')

def H_SummaryDetails():
    if not wd.db("Monthly_History_TM12").empty:
        wd.open("Monthly_History_TM12")
        wd.join("TM12_DTLS_join", right=wd.db("TM12_SUMM"), on=["UTCID"], how="inner")
        wd.extract("TM12_DTLS", filter="UTCID != '' and TRANSACTION_TYPE == 'SALE'")

def K_ExportDatabase():
    if not wd.db("TM12_DTLS").empty:
        wd.open("TM12_DTLS")
        wd.extract("TM12_DTLS_export", cols=["UTCID","ACCT_NO","TRUST_CODE","CUSTOMER_BRANCH","BRANCH_NAME","HOLDER_NAME","RECEIPT_NUMBER","TRANSACTION_BRANCH","BRANCH_DESCRIPTION","TRANSACTION_TYPE","TRAN_DATE","TRANSACTION_AMOUNT","UNITS","UNIT_PRICE","POST_DATE","OCCUPATION","HOLDER_TYPE","RISK_RATING","RATING_SOURCE","RUN_DATE","PAYMENT_TYPE"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM12_DTLS.mdb'))

def L_ExportSumm():
    if not wd.db("TM12_SUMM").empty:
        wd.open("TM12_SUMM")
        wd.extract("TM12_SUMM_export", cols=["UTCID","CUSTOMER_BRANCH","BRANCH_NAME","STAFF","HOLDER_TYPE","HOLDER_NAME","OCCUPATION","RATING_SOURCE","RISK_RATING","RUN_DATE","TRANSACTION_DAYS","NO_OF_TRANSACTIONS","TRANSACTION_SUMMARY"])
        wd.exportMDB2(filename=join(PARENT_DIR, 'Reports', 'TM12_SUMM.mdb'))

def P_Cleanup():
    if wd.tblName:
        wd.close()
    wd.delete("Deposits")
    wd.delete("Monthly_History_TM12")
    wd.delete("Monthly_History12")
    wd.delete("TM12_SUMM_INT")
    wd.delete("TM12_SUMM_INT2")
    wd.delete("TM12_SUMM")
    wd.delete("TM12_DTLS")

if __name__ == "__main__":
    wd = DataAnalytics()

    A_JoinCustTransactions()
    B_SummWeeklyTransDay()
    C_SummWeeklyTransExclude()
    F_AboveDollarThreshold()
    G_Rename_SummFields()
    H_SummaryDetails()
    K_ExportDatabase()
    L_ExportSumm()
    P_Cleanup()
from DataAnalytics import DataAnalytics
wd = DataAnalytics()

#Summarize History to get Customer Average by Transaction Type

def ASummHist_Average():
	if not wd.open("Tran_Hist_Average").empty:
		wd.open("Tran_Hist_Average")
		wd.summBy("Summ_Hist_Average", ['UTCID', 'TRANSACTION_TYPE'], agg_funcs={key: ['sum', 'mean'] if key != "UTCID" else ['count'] for key in ['TRANSACTION_AMOUNT', 'UTCID']})
		wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})
		#Create Join Key in Todays Database

def CCreateSumm_Today():
	if not wd.open("Daily_Transactions_Today").empty:
		wd.open("Daily_Transactions_Today")
		wd.summBy("Summ_Tran_Today", ['UTCID', 'TRANSACTION_TYPE', 'POST_DATE'], agg_funcs={key: ['sum'] if key != "UTCID" else ['count'] for key in ['TRANSACTION_AMOUNT', 'UTCID']})
		wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})
		wd.join("Summ_Tran_Today", right=wd.db("Summ_Tran_Today_summ")[['CUSTOMER_BRANCH', 'HOLDER_TYPE', 'HOLDER_NAME', 'OCCUPATION', 'RATING_SOURCE', 'RISK_RATING', 'BRANCH_NAME']], how="left")
		
from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def ASummHist_Average():
	wd.open("Tran_Hist_Average")
	wd.summBy("Summ_Hist_Average", ['UTCID', 'TRANSACTION_TYPE'], agg_funcs={key: ['sum'] if key != "UTCID" else ['sum', 'count'] for key in ['TRANSACTION_AMOUNT']})
	wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})
	def ASummHist_Average():
	if not wd.open("Tran_Hist_Average").empty:
		wd.open("Tran_Hist_Average")
		wd.summBy("Summ_Hist_Average", ['UTCID', 'TRANSACTION_TYPE', 'UTCID', 'TRANSACTION_TYPE'], agg_funcs={key: ['sum'] if key != "UTCID" else ['sum', 'count'] for key in ['TRANSACTION_AMOUNT', 'TRANSACTION_AMOUNT']})
		wd.renameCol(columns={"UTCID_count": "NO_OF_RECS"})
		
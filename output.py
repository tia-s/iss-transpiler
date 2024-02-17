from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def Get_Multiple_Cheques_deposited():
	#d

	exclude: {'fields': 'all', 'key': 'CUST_ID', 'diff_field': 'INSTRMNT_NUM', 'db_name': 'Multiple_Cheques_deposited_to_same_account.IMD', 'virt_db': False, 'perf_task': ''}
	
from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def BExtResult_Summ():
	if not wd.open("TM5_Summ_INT").empty:
		wd.open("TM5_Summ_INT")
		
from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def SCleanup():
	wd.delete("Master_Funds")
	wd.delete("Master_History")
	wd.delete("DPUTR04N")
	wd.delete("DPURF031")
	wd.delete("DPUTR011")
	wd.delete("DPUTR021")
	wd.delete("DPUTR081")
	wd.delete("Master_History_File")
	wd.delete("Transaction_History_INT")
	wd.delete("Transaction_History")
	wd.delete("History_Transaction_Hist")
	wd.delete("Risk_Ratings-")
	wd.delete("Risk Ratings and PORTAL")
	wd.delete("History_Transaction_Hist_Risk")
	
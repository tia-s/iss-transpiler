from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def EJoinExtResult_Summ():
	summ: {'Add to Summarize': ['UTCID', 'POST_DATE'], 'Add to Inc': ['HOLDER_TYPE', 'HOLDER_NAME', 'CUSTOMER_BRANCH', 'OCCUPATION', 'RATING_SOURCE', 'RISK_RATING', 'BRANCH_NAME'], 'Add to Total': ['TRANSACTION_AMOUNT'], 'Criteria': '', 'dbname': 'TM5_Summ1.IMD', 'Output DB Name': '', 'create_percnt': FALSE, 'stats': ['SM_SUM'], 'Perform Task': ''}
	def FRename_ResultFields():
		table: {'name': 'NO_OF_TRANSACTIONS', 'description': 'Number of records found for this key value', 'field_type': WI_NUM_FIELD, 'equation': '', 'decimals': 0, 'replace': 'NO_OF_RECS1'}
		table: {'name': 'TRANSACTION_SUMMARY', 'description': '', 'field_type': WI_NUM_FIELD, 'equation': '', 'decimals': 2, 'replace': 'TRANSACTION_AMOUNT_SUM'}
		
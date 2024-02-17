from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def G_Rename_SummFields():
	table: {'name': 'NO_OF_TRANSACTIONS', 'description': '', 'field_type': WI_NUM_FIELD, 'equation': '', 'decimals': 0, 'replace': 'NO_OF_RECS_SUM'}
	table: {'name': 'TRANSACTION_DAYS', 'description': '', 'field_type': WI_NUM_FIELD, 'equation': '', 'decimals': 0, 'replace': 'NO_OF_RECS1'}
	table: {'name': 'TRANSACTION_SUMMARY', 'description': '', 'field_type': WI_NUM_FIELD, 'equation': '', 'decimals': 2, 'replace': 'TRANSACTION_AMOUNT_SUM_SUM'}
	table: {'name': 'STAFF', 'description': '', 'field_type': WI_VIRT_CHAR, 'equation': '', 'length': 50}
	
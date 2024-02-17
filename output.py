from DataAnalytics import DataAnalytics
wd = DataAnalytics()

# File: Join Databases

def JoinDatabase():
	join: {'file_to_join': 'PUB.mb.IMD', 'p_fields': 'all', 's_fields': 'all', 'match_keys': ['GROUP_ID_AVERAGE', 'CU', 'A'], 'create_virtual_db': False, 'db_name': 'Join Databases2.IMD', 'perform_task': 'WI_JOIN_NOC_PRI_MATCH'}
	
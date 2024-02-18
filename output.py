from DataAnalytics import DataAnalytics
wd = DataAnalytics()

def BAJoinCustomers():
wd.join("Registrations", right=wd.db("REPORTSVR.VDPUCID01")[['"ADDRESS1"', '"ADDRESS2"', '"ADDRESS3"', '"BIRTHDAT"', '"HOLDRTYP"', '"FIRSTNAME"', '"SURNAME"']], how=WI_JOIN_ALL_IN_PRIM, on=['"UTCID"', '"UTCID"', '"A"'])

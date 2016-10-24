import pyodbc
import os, sys, json, base64
from types import *
from openpyxl import Workbook
import traceback

kwargs = dict()
kwargs['driver'] = '{SQL Server}'
kwargs['server'] = 'Reno-fis-sql2'
kwargs['database'] = 'ABM_Reno_GIS'
# kwargs['uid'] = 'RENOAIRPORT\\AroraTeam'
# kwargs['pwd'] = r'@R0r@G1$'

connGIS = pyodbc.connect(**kwargs)

kwargs['database'] = 'ABM_Reno_Prod'
connPROD = pyodbc.connect(**kwargs)

kwargs['database'] = 'ABM_Reno_Test'
connTEST = pyodbc.connect(**kwargs)


def queryConnection(database, connection):
    """take in the _mssql connection and write out geometries"""
    # query each connection database
    cursor = connection.cursor()
    cursor.execute('SELECT * FROM sys.tables')
    gis_tables = []
    rows = cursor.fetchall()
    for x in rows:
        gis_tables.append(x[0])

    for t in gis_tables:
        # for each table print out all of the rows
        cursor.execute('SELECT * FROM {}'.format(t))
        rows = cursor.fetchall()
        if rows:
            gis_wb = Workbook()
            ws = gis_wb.active
            headers = []
            for column in cursor.description:
                headers.append(column[0])
            ws.append(headers)

            for row in rows:
                values = []

                for v in row:
                    if type(v) == bytearray:
                        values.append(base64.b64encode(v))
                        # values.append(base64.b32encode(v))
                    elif type(v) is StringType:
                        values.append(v.encode('utf-8'))
                    elif type(v) is IntType:
                        values.append(str(int(v)).encode('utf-8'))
                    elif type(v) is LongType:
                        values.append(str(long(v)).encode('utf-8'))
                    elif type(v) is FloatType:
                        values.append(str(float(v)).encode('utf-8'))
                    elif type(v) is NoneType:
                        values.append("Null".encode('utf-8'))
                    elif type(v) is UnicodeType:
                        values.append(v.encode('utf-8'))
                    elif type(v) is BooleanType:
                        values.append("True".encode('utf-8'))
                try:
                    ws.append(values)
                except:
                    exc_type, exc_value, exc_traceback = sys.exc_info()
                    print("{} :: {}".format(traceback.print_tb(exc_traceback, limit=1, file=sys.stdout),
                                            traceback.print_exception(exc_type, exc_value, exc_traceback,
                                                                      limit=2, file=sys.stdout)))

            dirpath = r"C:\Users\arorateam\{}".format(database)
            if not os.path.exists(dirpath):
                os.mkdir(dirpath)
            outpath = r"{}\{}.xlsx".format(dirpath, t)
            if os.path.exists(outpath):
                os.remove(outpath)
            gis_wb.save(outpath)

for k, value in {
    # "ABM_Reno_GIS": connGIS,
    "ABM_Reno_Prod": connPROD,
    # "ABM_Reno_Test": connTEST
}.iteritems():
    try:
        queryConnection(k, value)
    except:
        traceback.print_exc(file=sys.stdout)

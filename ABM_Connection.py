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
                geo_flag = 0
                for v in row:
                    if type(v) == bytearray:
                        values.append(base64.b64encode(v))
                        # values.append(base64.b32encode(v))
                        geo_flag = 1
                    elif type(v) == unicode:
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

                if geo_flag and len(values):
                    ws.append(values)
                # ws.append(values)
                print values

            dirpath = r"C:\Users\arorateam\{}".format(database)
            if not os.path.exists(dirpath):
                os.mkdir(dirpath)
            outpath = r"{}\{}.xlsx".format(dirpath, t)
            if os.path.exists(outpath):
                os.remove(outpath)
            gis_wb.save(outpath)

for k, value in {"ABM_Reno_GIS": connGIS, "ABM_Reno_Prod": connPROD, "ABM_Reno_Test": connTEST}.iteritems():
    try:
        queryConnection(k, value)
    except:
        traceback.print_exc(file=sys.stdout)

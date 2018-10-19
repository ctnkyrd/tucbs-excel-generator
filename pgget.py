# -*- coding: utf-8 -*-
import psycopg2 as pg

class Connection:
    def __init__(self):
        self.cs = "dbname=%s user=%s password=%s host=%s port=%s" % ('tucbsdata','postgres','kalman','localhost','5432')

    def executeSql(self, query):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute(query)
        data = cur.fetchall()
        return data

    def getsinglekoddata(self, tableName, valueColumn, where = ""):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (valueColumn, tableName, where)+" limit 1")
        data = cur.fetchone()
        conn.close()
        return data[0]
    
    def getRowOfData(self, tableName, valueColumn, where = ""):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (valueColumn, tableName, where)+" limit 1")
        data = cur.fetchone()
        conn.close()
        return data

    def getSingledataByOid(self, tableName, valueColumn,oid, primaryKey="objectid"):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (valueColumn, tableName, primaryKey + "=" + str(oid)))
        data = cur.fetchone()
        conn.close()
        return data[0]

    def getlistofdata(self, tableName, columns='*', where=''):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (columns,tableName, where))
        data = cur.fetchall()
        conn.close()
        return data
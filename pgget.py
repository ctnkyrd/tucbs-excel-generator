# -*- coding: utf-8 -*-
import psycopg2 as pg

class Connection:
    def __init__(self):
        self.cs = "dbname=%s user=%s password=%s host=%s port=%s" % ('tucbsdata','postgres','kalman','localhost','5432')

    def getsinglekoddata(self, tableName, valueColumn, where = ""):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (valueColumn, tableName, where)+" limit 1")
        data = cur.fetchone()
        return data[0]

    def getlistofdata(self, tableName, columns='*', where=''):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (columns,tableName, where))
        data = cur.fetchall()
        return data
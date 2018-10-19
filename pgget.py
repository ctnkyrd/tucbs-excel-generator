# -*- coding: utf-8 -*-
import psycopg2 as pg

class Connection:
    def __init__(self):
        self.cs = "dbname=%s user=%s password=%s host=%s port=%s" % ('tucbsdata','postgres','Ankara123','192.168.30.136','5432')

    def getsinglekoddata(tableName, valueColumn, where = ""):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (valueColumn, tableName, where)+" limit 1")
        data = cur.fetchone()
        return data[0]

    def getlistofdata(tableName, columns='*', where=''):
        conn = pg.connect(self.cs)
        cur = conn.cursor()
        cur.execute("select %s from %s where %s" % (tableName, columns, where))
        data = cur.fetchall()
        return data
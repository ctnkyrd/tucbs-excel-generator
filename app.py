# -*- coding: utf-8 -*-

import psycopg2 as pg, datetime
from pgget import Connection as c
from kurum import Kurum

cs = "dbname=%s user=%s password=%s host=%s port=%s" % ('tucbsdata','postgres','Ankara123','192.168.30.136','5432')
conn = pg.connect(cs)
cur = conn.cursor()
cur.execute("select * from kurum where analiz_tamamlandi_first is true")
kurum = cur.fetchall()

for i in kurum:
    a = Kurum(i[0], i[3], i[8], i[2], i[1])
    print a.bakanlik
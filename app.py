# -*- coding: utf-8 -*-

import psycopg2 as pg, datetime
from pgget import Connection
from kurum import Kurum

cnn = Connection()

kurum = cnn.getlistofdata('kurum','*','analiz_tamamlandi_first is true')

for i in kurum:
    a = Kurum(i[0], i[3], i[8], i[2], i[1])
    print a.bakanlik
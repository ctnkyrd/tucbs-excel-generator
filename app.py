# -*- coding: utf-8 -*-
from pgget import Connection
from kurum import Kurum
from excel_format.cv import CografiVeriFormu

# create connection
cnn = Connection()

kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')

for i in kurum:
    a = Kurum(i[0])
    print a.bakanlik, a.adi, 'Started!'
    c = CografiVeriFormu(a.bakanlik, a.adi, a.birim)
    c.createExcelFile()
# -*- coding: utf-8 -*-
import datetime

print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]') + "Importing Modules"

from pgget import Connection
from kurum import Kurum
from excel_format.cv import CografiVeriFormu
from excel_format.cv_dict import dict_veri_katmani

# create connection
cnn = Connection()
kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')

cvdict = dict_veri_katmani
for i in kurum:
    newKurum = Kurum(i[0])
    print newKurum.bakanlik +"-->"+ newKurum.adi, 'Started!'
    ek2_oid = newKurum.ek2_oid
    
    kurumKatmanalri = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani','*','geodurum is true and ek_2='+str(ek2_oid))
    for katman in kurumKatmanalri:
        newKurum.add_veri_katmani(katman)
        print katman[cvdict['katman_adi']]

    # c = CografiVeriFormu(a.bakanlik, a.adi, a.birim)
    # c.createExcelFile()
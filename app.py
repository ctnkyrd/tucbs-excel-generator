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
        #pylint: disable-msg=too-many-arguments
        cvf = CografiVeriFormu(newKurum.bakanlik, newKurum.adi, newKurum.birim, katman[cvdict['tucbs_katmani']], katman[cvdict['katman_adi']], 
                            katman[cvdict['katman_durumu']], katman[cvdict['tucbs_uygunluk']], katman[cvdict['veri_turu']], katman[cvdict['veri_tipi']],
                            katman[cvdict['veri_adedi']], katman[cvdict['veri_formati']], katman[cvdict['projeksiyon']], katman[cvdict['datum']], 
                            katman[cvdict['olcek_duzey']], katman[cvdict['veri_guncelleme_periyod']], katman[cvdict['son_veri_guncelleme_tarih']], 
                            katman[cvdict['veri_envanteri_aciklama']], katman[cvdict['tucbs_tema_harici']], katman[cvdict['inspire_katmani']],
                            katman[cvdict['inspire_uygunluk']], katman[cvdict['katman_aciklama']], katman[cvdict['tesim_alindi']], katman[cvdict['teslim_formati']], katman[cvdict['teslim_alinan_veri_sayisi']])
        print cvf.katman_adi
        try:
            cvf.createExcelFile()
        except BaseException as be:
            print be.message
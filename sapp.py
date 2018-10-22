# -*- coding: utf-8 -*-
import datetime

print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]') + "Importing Modules"

from pgget import Connection
from kurum import Kurum
from excel_format.cv import CografiVeriFormu
from excel_format.cv_dict import dict_veri_katmani
from excel_format.mv import MetaveriFormu
from excel_format.mv_dict import dict_metaveri_katmani

# create connection
cnn = Connection()
kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')
counter=0
cvdict = dict_veri_katmani
mvdict = dict_metaveri_katmani
for i in kurum:
    
    newKurum = Kurum(i[0])
    print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]'), unicode(newKurum.adi)
    ek2_oid = newKurum.ek2_oid
    
    kurumKatmanalri = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani','*','geodurum is true and ek_2='+str(ek2_oid))
    for katman in kurumKatmanalri:
        
        #pylint: disable-msg=too-many-arguments
        mvf = MetaveriFormu(katman[mvdict['katman_adi']],katman[mvdict['mv_metaveri_var']],katman[mvdict['mv_standart']],
                            katman[mvdict['mv_yayinlaniyor']],katman[mvdict['mv_cbs_gm_paylasim_var']],katman[mvdict['metaveri_aciklama']],
                            newKurum.adi,katman[cvdict['tucbs_katmani']],katman[cvdict['inspire_katmani']])
        
        try:
            counter += 1
            mvf.createExcelFile()

        except BaseException as be:
            print be
# -*- coding: utf-8 -*-
import datetime

print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]') + " Started"
import sys
from pgget import Connection
from kurum import Kurum
from excel_format.cv import CografiVeriFormu
from excel_format.cv_dict import dict_veri_katmani
from excel_format.mv import MetaveriFormu
from excel_format.mv_dict import dict_metaveri_katmani

# create connection
cnn = Connection()
kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')

cvdict = dict_veri_katmani
mvdict = dict_metaveri_katmani
for i in kurum:
    sys.stdout.flush()
    counter=0
    newKurum = Kurum(i[0])
    # print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]'), unicode(newKurum.adi)
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
                            katman[cvdict['inspire_uygunluk']], katman[cvdict['katman_aciklama']], katman[cvdict['tesim_alindi']], katman[cvdict['teslim_formati']], 
                            katman[cvdict['teslim_alinan_veri_sayisi']],  katman[cvdict['vk_amac']], katman[cvdict['vk_kullanim']], katman[cvdict['vk_kokeni']], 
                            katman[cvdict['vk_copleteness_fazlalik']], katman[cvdict['vk_fazlalik_yeni']], katman[cvdict['vk_eksizlik_yeni']], 
                            katman[cvdict['vk_lc_kavramsal_tutarlilik']], katman[cvdict['vk_kavramsal_yeni']], katman[cvdict['vk_tanim_kumesi_yeni']], 
                            katman[cvdict['vk_format_tutarlilik_yeni']], katman[cvdict['vk_topoloji_tutarlilik_yeni']], katman[cvdict['vk_pa_mutlak_dogruluk']], 
                            katman[cvdict['vk_konumsal_mutlak_dogruluk_yeni']], katman[cvdict['vk_konumsal_bagil_dogruluk_yeni']], katman[cvdict['vk_konumsal_raster_veri_konum_yeni']], 
                            katman[cvdict['vk_ta_ilgili_zamandaki_dogruluk']], katman[cvdict['vk_zamansal_ilgili_yeni']], katman[cvdict['vk_zamansal_tutarlilik_yeni']], 
                            katman[cvdict['vk_zamansal_gecerlilik_yeni']], katman[cvdict['vk_tema_siniflandirma_dogrulugu']], katman[cvdict['vk_tematik_siniflandirma_yeni']], 
                            katman[cvdict['vk_tematik_nicel_yeni']], katman[cvdict['vk_tematik_nicel_olmayan_yeni']], katman[cvdict['vk_aciklama']])
        
        mvf = MetaveriFormu(katman[mvdict['katman_adi']],katman[mvdict['mv_metaveri_var']],katman[mvdict['mv_standart']],
                            katman[mvdict['mv_yayinlaniyor']],katman[mvdict['mv_cbs_gm_paylasim_var']],katman[mvdict['metaveri_aciklama']],
                            newKurum.adi,katman[cvdict['tucbs_katmani']],katman[cvdict['inspire_katmani']])
        
        cvf.createExcelFile()
        mvf.createExcelFile()
        try:
            counter += 1
            sys.stdout.write(unicode(newKurum.adi)+u"Katman Sayısı: %d   \r" % (counter))
            sys.stdout.flush()
            # print str(counter) +"-->"+ cvf.katman_adi.decode('utf-8')
            # excel created here
            cvf.createExcelFile()
        except BaseException as be:
            print be
    print unicode(newKurum.adi), u"--> Tamamlandı", u" Toplam Katman: "+ str(counter)
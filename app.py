# -*- coding: utf-8 -*-
import datetime

print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]') + "Started"
import sys, os
from pgget import Connection
from kurum import Kurum
from excel_format.cv import CografiVeriFormu
from excel_format.dicts import *
from excel_format.mv import MetaveriFormu
from excel_format.dyag import DonanimYazilimFormu
from excel_format.ob import OrganizasyonBirimiFormu
from excel_format.ma import MevzuatAnalizFormu



# create connection
cnn = Connection()
kurum = cnn.getlistofdata('kurum','objectid','analiz_tamamlandi_first is true')

cvdict = dict_veri_katmani
mvdict = dict_metaveri_katmani
dyagdict = dict_donanim_yazilim
orbidict = dict_organizasyon_birimleri
madict = dict_mevzuat_analizi
for i in kurum:
    sys.stdout.flush()
    counter=0
    newKurum = Kurum(i[0])
    # print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]'), unicode(newKurum.adi)
    ek2_oid = newKurum.ek2_oid
    kurumKatmanalri = cnn.getlistofdata('x_ek_2_tucbs_veri_katmani','*','geodurum is true and ek_2='+str(ek2_oid))
    kurumDonanimYazilimAgGuvenlik = cnn.getlistofdata('ek_donanim_yazilim_ag_guvenlik', '*', 'geodurum is true and kurum='+str(newKurum.oid))
    organizasyonBirimleri = cnn.getlistofdata('ek_cbs_organizasyon_birimleri_analizi', '*', 'geodurum is true and kurum='+str(newKurum.oid))
    mevzuatanalizi = cnn.getlistofdata('ek_mevzuat_analizi', '*', 'geodurum is true and kurum='+str(newKurum.oid))
    

    ma_counter = 0
    for ma in mevzuatanalizi:
        maf = MevzuatAnalizFormu(newKurum.bakanlik, newKurum.adi, newKurum.k_adi, newKurum.birim, ma_counter, ma[madict['objectid']], ma[madict['mevzuat_kisitlama']],
                                ma[madict['veri_paylasmama_sebep']])

        ma_counter += 1

        maf.createExcelFile()

    # birden fazla donanım yazılım formu olması durumu için oluşturuldu
    dyag_counter = 0
    for dyag in kurumDonanimYazilimAgGuvenlik:
        dyagvf =  DonanimYazilimFormu(newKurum.bakanlik, newKurum.adi, newKurum.birim, dyag_counter, dyag[dyagdict['sunucu_yeterli']],dyag[dyagdict['sunucu_yetersiz_aciklama']],
                                        dyag[dyagdict['kamunet_agina_bagli']],dyag[dyagdict['kamunet_agina_bagli_degil_aciklama']],
                                        dyag[dyagdict['ipsecvpn_uygun']],dyag[dyagdict['ipsecvpn_uygunsuz_aciklama']],dyag[dyagdict['ipsecvpn_bagli']], newKurum.k_adi)
        dyag_counter += 1

        dyagvf.createExcelFile()

    # birden fazla organizasyon birimi formu olması durumu için oluşturuldu
    orbi_counter = 0
    for orbi in organizasyonBirimleri:
        obf = OrganizasyonBirimiFormu(newKurum.bakanlik, newKurum.adi, newKurum.k_adi,orbi_counter, orbi[orbidict['cbs_birimi_var']],orbi[orbidict['cbs_yeni_birim_gerekli']],
                                    orbi[orbidict['veri_uretim_var']],orbi[orbidict['tasra_teskilati_var']],orbi[orbidict['sadece_sistem_idamesi_var']],orbi[orbidict['ihtiyaclar_sizin_onayinizdan_geciyor']],
                                    orbi[orbidict['personel_yeterli']],orbi[orbidict['personel_yetersizlik_kriterleri']],orbi[orbidict['personel_yetersizlik_oneri']],orbi[orbidict['cbs_farkindaligi_var']],
                                    orbi[orbidict['sorunlar']],orbi[orbidict['cbs_birim_adi']],orbi[orbidict['yapilanma_olcegi']],orbi[orbidict['kurum_semasindaki_yeri']],orbi[orbidict['tasra_teskilat_yapilanmasi']],
                                    orbi[orbidict['cbs_yeni_birim_gorusler']])
        orbi_counter += 1
        obf.createExcelFile()
    
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
                            katman[cvdict['vk_tematik_nicel_yeni']], katman[cvdict['vk_tematik_nicel_olmayan_yeni']], katman[cvdict['vk_aciklama']], newKurum.k_adi)
        
        mvf = MetaveriFormu(katman[mvdict['katman_adi']],katman[mvdict['mv_metaveri_var']],katman[mvdict['mv_standart']],
                            katman[mvdict['mv_yayinlaniyor']],katman[mvdict['mv_cbs_gm_paylasim_var']],katman[mvdict['metaveri_aciklama']],
                            newKurum.adi,katman[cvdict['tucbs_katmani']],katman[cvdict['inspire_katmani']], newKurum.k_adi)
        try:
            counter += 1
            cvf.createExcelFile()
            mvf.createExcelFile()
            sys.stdout.write(unicode(newKurum.adi)+u"Katman Sayısı: %d   \r" % (counter))
            sys.stdout.flush()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
    print datetime.datetime.now().strftime('[%Y-%m-%d][%H:%M:%S]') + unicode(newKurum.adi), u"--> Tamamlandı", u" Toplam Katman: "+ str(counter)
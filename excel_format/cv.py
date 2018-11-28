# -*- coding: utf-8 -*-
import xlsxwriter, os, sys
from pgget import Connection
cnn = Connection()

class CografiVeriFormu:
    def __init__(self, bakanlik, adi, birim, tucbs_katmani, katman_adi, katman_durumu, tucbs_uygunluk, veri_turu, veri_tipi, veri_adedi, veri_formati, projeksiyon,
                    datum, olcek_duzey, veri_guncelleme_periyod, son_veri_guncelleme_tarih, veri_envanteri_aciklama, tucbs_tema_harici, inspire_katmani, inspire_uygunluk,
                    katman_aciklama, tesim_alindi, teslim_formati, teslim_alinan_veri_sayisi, vk_amac,vk_kullanim,vk_kokeni,vk_copleteness_fazlalik, vk_fazlalik_yeni,
                    veri_eksiklik_additional,vk_lc_kavramsal_tutarlilik,vk_kavramsal_yeni,vk_tanim_kumesi_yeni,vk_format_tutarlilik_yeni,vk_topoloji_tutarlilik_yeni,
                    vk_pa_mutlak_dogruluk,vk_konumsal_mutlak_dogruluk_yeni,vk_konumsal_bagil_dogruluk_yeni,vk_konumsal_raster_veri_konum_yeni,
                    vk_ta_ilgili_zamandaki_dogruluk,vk_zamansal_ilgili_yeni,vk_zamansal_tutarlilik_yeni,vk_zamansal_gecerlilik_yeni,vk_tema_siniflandirma_dogrulugu,
                    vk_tematik_siniflandirma_yeni,vk_tematik_nicel_yeni,vk_tematik_nicel_olmayan_yeni, vk_aciklama, k_adi, geom_yeni, ve_duzey):
        self.bakanlik = bakanlik
        self.adi = adi.rstrip()
        self.birim = birim
        self.k_adi = k_adi
        if geom_yeni is not None:
            self.geom_yeni = geom_yeni
        else:
            self.geom_yeni = u''
        if ve_duzey is not None:
            self.ve_duzey = cnn.getsinglekoddata('kod_veri_envanteri_duzey', 'kod', 'objectid='+str(ve_duzey))
        else:
            self.ve_duzey = u''
        # tucbs temasinin kod tablosundan çekilmesi
        if tucbs_katmani is not None:
            self.tucbs_veri_temasi = cnn.getsinglekoddata('kod_tucbs_tema', 'tema_adi', 'objectid='+str(tucbs_katmani))
        else:
            self.tucbs_veri_temasi = None
        self.katman_adi = katman_adi.rstrip()
        self.katman_durumu = katman_durumu
        self.tucbs_uygunluk = tucbs_uygunluk

        # veri envanteri
        if veri_turu is not None:
            self.veri_turu = cnn.getsinglekoddata('kod_ek_2_veri_turu', 'kod', 'objectid='+str(veri_turu))
        else:
            self.veri_turu = None
        if veri_tipi is not None:
            self.veri_tipi = cnn.getsinglekoddata('kod_ek_2_veri_tipi', 'kod', 'objectid='+str(veri_tipi))
        else:
            self.veri_tipi = None 
        self.veri_adedi = veri_adedi
        if veri_formati is not None:
            self.veri_formati = cnn.getsinglekoddata('kod_ek_2_veri_formati', 'kod', 'objectid='+str(veri_formati))
        else:
            self.veri_formati = None
        if projeksiyon is not None:
            self.projeksiyon = cnn.getsinglekoddata('kod_ek_2_projeksiyon', 'kod', 'objectid='+str(projeksiyon))
        else:
            self.projeksiyon = None
        if datum is not None:
            self.datum = cnn.getsinglekoddata('kod_ek_2_datum', 'kod', 'objectid='+str(datum))
        else:
            self.datum = None
        self.olcek_duzey = olcek_duzey
        self.veri_guncelleme_periyod = veri_guncelleme_periyod
        self.son_veri_guncelleme_tarih = son_veri_guncelleme_tarih
        self.veri_envanteri_aciklama = veri_envanteri_aciklama
        self.tucbs_tema_harici = tucbs_tema_harici
        
        if inspire_katmani is not None:
            self.inspire_katmani = cnn.getsinglekoddata('kod_inspire_tema', 'tema_adi', 'objectid='+str(inspire_katmani))
        else:
            self.inspire_katmani = None
        self.inspire_uygunluk = inspire_uygunluk

        if self.tucbs_veri_temasi is None and self.inspire_katmani is not None:
            self.tucbs_tema_harici = True
        elif self.tucbs_veri_temasi is not None and self.inspire_katmani is None:
            self.tucbs_tema_harici = False
        else:
            self.tucbs_tema_harici = False

        self.katman_aciklama = katman_aciklama        
        self.teslim_alindi = tesim_alindi

        if teslim_formati is not None:
            self.teslim_formati = cnn.getsinglekoddata('kod_ek_2_veri_formati', 'kod', 'objectid='+str(teslim_formati))
        else:
            self.teslim_formati = None
        self.teslim_alinan_veri_sayisi = teslim_alinan_veri_sayisi

        # veri kalitesi columns get
        self.vk_amac = vk_amac
        self.vk_kullanim = vk_kullanim
        self.vk_kokeni = vk_kokeni
        self.vk_copleteness_fazlalik = vk_copleteness_fazlalik

        if vk_fazlalik_yeni is not None:
            self.vk_fazlalik_yeni = cnn.getsinglekoddata('kod_ek_2_fazlalik', 'kod', 'objectid='+str(vk_fazlalik_yeni))
        else:
            self.vk_fazlalik_yeni = None
        if veri_eksiklik_additional is not None:
            self.veri_eksiklik_additional = veri_eksiklik_additional
            #self.vk_eksizlik_yeni = cnn.getsinglekoddata('kod_ek_2_verinin_eksiksizligi', 'kod', 'objectid='+str(vk_eksizlik_yeni))
        else:
            self.veri_eksiklik_additional = veri_eksiklik_additional
            #self.vk_eksizlik_yeni = None
        self.vk_lc_kavramsal_tutarlilik = vk_lc_kavramsal_tutarlilik
        if vk_kavramsal_yeni is not None:
            self.vk_kavramsal_yeni = cnn.getsinglekoddata('kod_ek_2_kavramsal_tutarlilik', 'kod', 'objectid='+str(vk_kavramsal_yeni))
        else:
            self.vk_kavramsal_yeni = None
        if vk_tanim_kumesi_yeni is not None:
            self.vk_tanim_kumesi_yeni = cnn.getsinglekoddata('kod_ek_2_tanim_kumesi_tutarlilik', 'kod', 'objectid='+str(vk_tanim_kumesi_yeni))
        else:
            self.vk_tanim_kumesi_yeni = None
        if vk_format_tutarlilik_yeni is not None:
            self.vk_format_tutarlilik_yeni = cnn.getsinglekoddata('kod_ek_2_format_tutarlilik', 'kod', 'objectid='+str(vk_format_tutarlilik_yeni))
        else:
            self.vk_format_tutarlilik_yeni = None
        if vk_topoloji_tutarlilik_yeni is not None:
            self.vk_topoloji_tutarlilik_yeni = cnn.getsinglekoddata('kod_ek_2_topoloji_tutarlilik', 'kod', 'objectid='+str(vk_topoloji_tutarlilik_yeni))
        else:
            self.vk_topoloji_tutarlilik_yeni = None

        self.vk_pa_mutlak_dogruluk = vk_pa_mutlak_dogruluk

        if vk_konumsal_mutlak_dogruluk_yeni is not None:
            self.vk_konumsal_mutlak_dogruluk_yeni = cnn.getsinglekoddata('kod_ek_2_konumsal_dogruluk', 'kod', 'objectid='+str(vk_konumsal_mutlak_dogruluk_yeni))
        else:
            self.vk_konumsal_mutlak_dogruluk_yeni = None
        
        if vk_konumsal_bagil_dogruluk_yeni is not None:
            self.vk_konumsal_bagil_dogruluk_yeni = cnn.getsinglekoddata('kod_ek_2_konumsal_dogruluk', 'kod', 'objectid='+str(vk_konumsal_bagil_dogruluk_yeni))
        else:
            self.vk_konumsal_bagil_dogruluk_yeni = None
        
        if vk_konumsal_raster_veri_konum_yeni is not None:
            self.vk_konumsal_raster_veri_konum_yeni = cnn.getsinglekoddata('kod_ek_2_konumsal_dogruluk', 'kod', 'objectid='+str(vk_konumsal_raster_veri_konum_yeni))
        else:
            self.vk_konumsal_raster_veri_konum_yeni = None

        self.vk_ta_ilgili_zamandaki_dogruluk = vk_ta_ilgili_zamandaki_dogruluk

        if vk_zamansal_ilgili_yeni is not None:
            self.vk_zamansal_ilgili_yeni = cnn.getsinglekoddata('kod_ek_2_ilgili_zamandaki_dogruluk', 'kod', 'objectid='+str(vk_zamansal_ilgili_yeni))
        else:
            self.vk_zamansal_ilgili_yeni = None
        
        if vk_zamansal_tutarlilik_yeni is not None:
            self.vk_zamansal_tutarlilik_yeni = cnn.getsinglekoddata('kod_ek_2_zamansal_tutarlilik', 'kod', 'objectid='+str(vk_zamansal_tutarlilik_yeni))
        else:
            self.vk_zamansal_tutarlilik_yeni = None

        if vk_zamansal_gecerlilik_yeni is not None:
            self.vk_zamansal_gecerlilik_yeni = cnn.getsinglekoddata('kod_ek_2_zamansal_gecerlilik', 'kod', 'objectid='+str(vk_zamansal_gecerlilik_yeni))
        else:
            self.vk_zamansal_gecerlilik_yeni = None
        self.vk_tema_siniflandirma_dogrulugu = vk_tema_siniflandirma_dogrulugu

        if vk_tematik_siniflandirma_yeni is not None:
            self.vk_tematik_siniflandirma_yeni = cnn.getsinglekoddata('kod_ek_2_siniflandirma_dogrulugu', 'kod', 'objectid='+str(vk_tematik_siniflandirma_yeni))
        else:
            self.vk_tematik_siniflandirma_yeni = None

        if vk_tematik_nicel_yeni is not None:
            self.vk_tematik_nicel_yeni = cnn.getsinglekoddata('kod_ek_2_nicel_oznitelik_bilgileri_dogruluk', 'kod', 'objectid='+str(vk_tematik_nicel_yeni))
        else:
            self.vk_tematik_nicel_yeni = None

        if vk_tematik_nicel_olmayan_yeni is not None:
            self.vk_tematik_nicel_olmayan_yeni = cnn.getsinglekoddata('kod_ek_2_nicel_olmayan_oznitelik_bilgileri_dogruluk', 'kod', 'objectid='+str(vk_tematik_nicel_olmayan_yeni))
        else:
            self.vk_tematik_nicel_olmayan_yeni = None
        self.vk_aciklama = vk_aciklama
        
    
    def createExcelFile(self):
        try:
            excelPath = "created_excels"+"\\"+self.k_adi.decode('utf-8')+"\\"+u"CVSPMV"
            excelName = u"CVAF.xlsx"
            temaName = u"Yok"
            katmanName = u"Yok"
            if self.tucbs_veri_temasi is not None:
                temaName = self.tucbs_veri_temasi.decode('utf-8')
            elif self.inspire_katmani is not None:
                temaName = self.inspire_katmani.decode('utf-8')
            
            if self.katman_adi is not None:
                katmanName = self.katman_adi.decode('utf-8')
                if '/' in katmanName:
                    katmanName = katmanName.replace('/', '_')
            else:
                katmanName = u"Yok"

            fullFolderPath = excelPath+"\\"+temaName.rstrip()+"\\"+katmanName
            if os.path.isdir(unicode(fullFolderPath)) is False:
                try:
                    os.makedirs(unicode(fullFolderPath))
                except Exception as e:
                    exc_type, exc_obj, exc_tb = sys.exc_info()
                    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    print(exc_type, fname, exc_tb.tb_lineno)
            
            wb = xlsxwriter.Workbook(fullFolderPath+"\\"+excelName)
            ws = wb.add_worksheet()
            ws.set_column('A:A', 5)
            ws.set_column('B:B', 5)
            ws.set_column('C:C', 10.29)
            ws.set_column('D:D', 11.86)
            ws.set_column('E:E', 5)
            ws.set_column('F:F', 1.14)
            ws.set_column('G:G', 5)
            ws.set_column('H:H', 5)
            ws.set_column('I:I', 5)
            ws.set_column('J:J', 5)
            ws.set_column('K:K', 5)
            ws.set_column('L:L', 5)
            ws.set_column('M:M', 4.29)
            ws.set_column('N:N', 5)
            ws.set_column('O:O', 3.57)
            ws.set_column('P:P', 7.14)
            ws.set_column('Q:Q', 1.19)
            ws.set_column('R:R', 14.71)
            ws.set_column('S:S', 19)
            ws.set_column('T:T', 18.57)
            ws.set_column('U:U', 16.23)
            ws.set_row(5, 21)
            ws.set_row(4, 5)
            ws.set_row(14, 5)
            ws.set_row(18, 30)
            ws.set_row(17, 21)
            ws.set_row(20, 24.75)
            ws.set_row(21, 14.25)
            ws.set_row(22, 100)
            ws.set_row(24, 5)
            ws.set_row(25, 21)
            ws.set_row(28, 5)
            ws.set_row(29, 21)
            ws.set_row(53, 20)
            ws.set_row(54, 20)
            ws.set_row(55, 5)
            ws.set_row(56, 21)

            merge_header_format = wb.add_format()
            merge_header_format.set_font_size(16)
            merge_header_format.set_bold()
            merge_header_format.set_border()
            merge_header_format.set_align('center')
            merge_header_format.set_align('vcenter')

            merge_header_format2 = wb.add_format()
            merge_header_format2.set_font_size(16)
            merge_header_format2.set_bold()
            merge_header_format2.set_border()

            merge_small_header = wb.add_format()
            merge_small_header.set_font_size(11)
            merge_small_header.set_bold()
            merge_small_header.set_border()
            merge_small_header.set_align('center')
            merge_small_header.set_align('vcenter')
            merge_small_header.set_text_wrap()

            merge_small_header_left = wb.add_format()
            merge_small_header_left.set_font_size(11)
            merge_small_header_left.set_bold()
            merge_small_header_left.set_border()
            merge_small_header_left.set_align('left')
            merge_small_header_left.set_align('vcenter')
            merge_small_header_left.set_text_wrap()

            xsmall_header_right = wb.add_format()
            xsmall_header_right.set_font_size(10)
            xsmall_header_right.set_border()
            xsmall_header_right.set_align('right')

            merge_small_header2 = wb.add_format()
            merge_small_header2.set_font_size(11)
            merge_small_header2.set_bold()
            merge_small_header2.set_border()
            merge_small_header2.set_text_wrap()
            merge_small_header2.set_align('vcenter')

            f_data_right = wb.add_format()
            f_data_right.set_bold()
            f_data_right.set_font_color('red')
            f_data_right.set_align('right')
            f_data_right.set_border()
            f_data_right.set_text_wrap()
            f_data_right.set_valign('vcenter')

            f_data_left = wb.add_format()
            f_data_left.set_bold()
            f_data_left.set_font_color('red')
            f_data_left.set_align('left')
            f_data_left.set_border()
            f_data_left.set_text_wrap()
            f_data_left.set_valign('vcenter')

            f_data_center = wb.add_format()
            f_data_center.set_bold()
            f_data_center.set_font_color('red')
            f_data_center.set_align('center')
            f_data_center.set_align('vcenter')
            f_data_center.set_border()
            f_data_center.set_text_wrap()

            f_red = wb.add_format()
            f_red.set_font_color('red')
            f_red.set_bold()
            f_red.set_border()

            f_border = wb.add_format()
            f_border.set_border()

            f_border_center = wb.add_format()
            f_border_center.set_border()
            f_border_center.set_align('center')
            f_border_center.set_align('vcenter')
            f_border_center.set_text_wrap()

            f_data_emty = wb.add_format()
            f_data_emty.set_bg_color('#C5C5C5')
            f_data_emty.set_border()
            f_data_emty.set_text_wrap()

            f_comment = wb.add_format()
            f_comment.set_border(),f_border_center
            f_comment.set_color('gray')
            f_comment.set_italic()
            f_comment.set_font_size(9)
            f_comment.set_align('right')
            f_comment.set_valign('vcenter')

            f_comment_left = wb.add_format()
            f_comment_left.set_border()
            f_comment_left.set_color('gray')
            f_comment_left.set_italic()
            f_comment_left.set_font_size(9)
            f_comment_left.set_valign('vcenter')

            ws.insert_image('A1', r"logo\csb.jpg", {'x_offset': 20,'y_offset': 7,'x_scale': 1.6})
            ws.merge_range('A1:D4','',merge_header_format)

            ws.merge_range('E1:Q4',u'Coğrafi Veri Analiz Formu', merge_header_format)

            ws.merge_range('R1:S1', u'Revizyon Numarası', merge_small_header)
            ws.merge_range('R2:S2', '',merge_small_header)
            ws.merge_range('R3:S3', u'Revizyon Tarihi', merge_small_header)
            ws.merge_range('R4:S4', '',merge_small_header)
            ws.merge_range('T1:U4', '',merge_small_header)
            ws.merge_range('A5:U5', '')
            ws.merge_range('A6:F6', u'Genel Bilgiler',merge_header_format2)
            ws.merge_range('G6:U6', u'TUCBS Analiz Portalı üzerinden doldurulacak',f_comment)
            ws.merge_range('G13:U14', '',merge_small_header)
            ws.merge_range('A15:U15', '')


            ws.insert_image('T1', r"logo\tucbs2.jpg", {'x_offset': 10,'y_offset': 3,'x_scale': 0.5,'y_scale': 0.5})

            ws.merge_range('A7:F7', u'Bakanlık', merge_small_header2)
            ws.merge_range('A8:F8', u'Genel Müdürlük / Belediye', merge_small_header2)
            ws.merge_range('A9:F9', u'Birim', merge_small_header2)
            ws.merge_range('A10:F10', u'TUCBS Coğrafi Veri Teması', merge_small_header2)
            ws.merge_range('A11:F14', u'TUCBS Veri Katmanları', merge_small_header)
            ws.merge_range('A16:F18', u'TUCBS Harici Üretilen Veri Katmanı Var Mı?', merge_small_header)
            # fill form
            ws.merge_range('G7:U7', self.bakanlik, f_data_right)
            ws.merge_range('G8:U8', self.adi, f_data_right)
            ws.merge_range('G9:U9', self.birim, f_data_right)
            
            ws.merge_range('G11:Q11', u'Veri Katman Adı', merge_small_header2)
            ws.merge_range('R11:S11', u'Veri Katman Durumu', merge_small_header)
            ws.merge_range('T11:U11', u'TUCBS Standartlarına Uygunluk', merge_small_header)
            ws.merge_range('G17:M17', u'Katmanın INSPIRE\' a Uygun Tema Adı', merge_small_header)
            ws.merge_range('G18:M18', u'Katman INSPIRE Standartlarına Uygun Mu?', merge_small_header)

            # veri envanteri
            ws.merge_range('A20:U20', u'', merge_small_header)
            ws.write_rich_string('A20', merge_header_format2, u'Veri Envanteri', f_comment_left, u'  (Her Bir Katman İçin Kuruma Sorulacaklar)',f_border)
            ws.merge_range('A21:C22', u'',f_border_center)
            ws.merge_range('D21:F22', u'',f_border_center)
            ws.merge_range('G21:I22', u'',f_border_center)
            ws.merge_range('J21:M22', u'',f_border_center)
            ws.merge_range('N21:P22', u'',f_border_center)
            ws.merge_range('Q21:R22', u'',f_border_center)
            ws.merge_range('S21:S22', u'',f_border_center)
            ws.merge_range('T21:T22', u'',f_border_center)
            ws.merge_range('U21:U22', u'',f_border_center)
            ws.write('A21',u'Katman Adı', f_border_center)
            ws.write_rich_string('D21', merge_small_header, u'Veri Türü', f_comment, u' (Dijital Veri / Basılı Veri)', f_border_center)
            ws.write_rich_string('G21', merge_small_header, u'Veri Tipi', f_comment, u' (Coğrafi Veri / Sözel Veri)', f_border_center)
            ws.write('J21', u'Veri Adetleri', merge_small_header)
            ws.write('N21', u'Veri Formatı', merge_small_header)
            ws.write('Q21', u'Projeksiyon/Datum Bilgisi', merge_small_header)
            ws.write('S21', u'Ölçek/Düzey/Çözünürlük', merge_small_header)
            ws.write_rich_string('T21', merge_small_header, u'Veri Güncelleme Durumu', f_comment, u' (Güncelleme Sıklığı)', f_border_center)
            ws.write('U21', u'Son Veri Güncelleme Tarihi', merge_small_header)

            # veri envanteri loaders
            ws.merge_range('A23:C23',self.katman_adi.decode('utf-8'), f_data_center)
            # veri turu
            if self.veri_turu is not None:
                ws.merge_range('D23:F23',self.veri_turu.decode('utf-8'), f_data_center)
            else:
                ws.merge_range('D23:F23', u'', f_data_emty)
            # veri tipi
            if self.veri_tipi is not None:
                ws.merge_range('G23:I23',self.veri_tipi.decode('utf-8'), f_data_center)
            else:
                ws.merge_range('G23:I23', u'', f_data_emty)
            # veri adetleri
            if self.veri_adedi is not None:
                ws.merge_range('J23:M23',self.veri_adedi.decode('utf-8'), f_data_center)
            else:
                ws.merge_range('J23:M23', u'', f_data_emty)
            # veri formati
            if self.veri_formati is not None:
                ws.merge_range('N23:P23', self.veri_formati.decode('utf-8') + "\n" + self.geom_yeni.decode('utf-8'), f_data_center)
            else:
                ws.merge_range('N23:P23', u'', f_data_emty)
            # projeksyion datum
            if self.projeksiyon is not None or self.datum is not None:
                ws.merge_range('Q23:R23',self.projeksiyon.decode('utf-8') + " " + self.datum.decode('utf-8'), f_data_center)
            else:
                ws.merge_range('Q23:R23', u'', f_data_emty)
            # ölçek düzey çözünürlük
            if self.olcek_duzey is not None:
                ws.write('S23',self.olcek_duzey.decode('utf-8')+'\n'+self.ve_duzey.decode('utf-8'), f_data_center)
            else:
                ws.write('S23', u'', f_data_emty)
            # veri güncelleme
            if self.veri_guncelleme_periyod is not None:
                ws.write('T23',self.veri_guncelleme_periyod.decode('utf-8'), f_data_center)
            else:
                ws.write('T23', u'', f_data_emty)
            # son veri güncelleme tarihi
            if self.son_veri_guncelleme_tarih is not None:
                ws.write('U23',self.son_veri_guncelleme_tarih.strftime('%d.%m.%Y'), f_data_center)
            else:
                ws.write('U23', u'', f_data_emty)

            # katman aciklama
            if self.veri_envanteri_aciklama is not None:
                h = 15*len(self.veri_envanteri_aciklama)/100
                ws.merge_range('A24:U24', u'', merge_small_header2)                
                ws.write('A24', self.veri_envanteri_aciklama.decode('utf-8'), f_data_center)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(23, h)
            else:
                ws.merge_range('A24:U24', u'', merge_small_header2)
                ws.set_row(23, 5)
            ws.merge_range('A25:U25', u'')

            # veri teslimi
            ws.merge_range('A26:U26', u'', merge_small_header)
            ws.write_rich_string('A26', merge_header_format2, u'Veri Teslimi', f_comment_left, u'  (Teslim alınan veriler için doldurulacaktır)',f_border)
            ws.merge_range('A27:F27', u'Veri Katmanı', merge_small_header)
            ws.merge_range('G27:L27', u'', merge_small_header)
            ws.write_rich_string('G27', merge_small_header, u'Veri Tipi', f_comment, u' (Coğrafi Veri / Sözel Veri)', f_border_center)
            ws.merge_range('M27:S27', u'Veri Formatı', merge_small_header)
            ws.merge_range('T27:U27', u'Veri Sayısı', merge_small_header)


            if self.teslim_alindi:
                ws.merge_range('A28:F28', self.katman_adi.decode('utf-8'),f_data_center)
                if self.veri_tipi is not None:
                     ws.merge_range('G28:L28', self.veri_tipi.decode('utf-8'),f_data_center)
                else:
                     ws.merge_range('G28:L28', u'', f_data_emty)
                
                if self.teslim_formati is not None:
                     ws.merge_range('M28:S28', self.teslim_formati.decode('utf-8'), f_data_center)
                else:
                     ws.merge_range('M28:S28', u'', f_data_emty)
                
                if self.teslim_alinan_veri_sayisi is not None:
                    ws.merge_range('T28:U28', self.teslim_alinan_veri_sayisi.decode('utf-8'), f_data_center)
                else:
                    ws.merge_range('T28:U28', u'', f_data_emty)
            else:
                ws.merge_range('A28:F28', u'', merge_small_header)
                ws.merge_range('G28:L28', u'', merge_small_header)
                ws.merge_range('M28:S28', u'', merge_small_header)
                ws.merge_range('T28:U28', u'', merge_small_header)

            ws.merge_range('G16:J16', u'', merge_small_header)
            ws.merge_range('K16:M16', u'', merge_small_header)
            ws.merge_range('N18:U18', u'', merge_small_header)

            if self.katman_aciklama is not None:
                h = 15*len(self.katman_aciklama)/100
                ws.merge_range('A19:U19', self.katman_aciklama.decode('utf-8'), f_data_left)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(18, h)
            else:
                ws.merge_range('A19:U19', u'')
                ws.set_row(18, 5)

            # veri kalitesi gereksinimleri
            ws.merge_range('A29:U29', u'')
            ws.merge_range('A30:U30', u'', merge_small_header2)
            ws.write_rich_string('A30', merge_header_format2, u'Veri Kalitesi Gereksinimleri', f_comment_left, u' (Her bir katman için sorulacaktır)', f_border)
            ws.merge_range('A31:G31', u'Amaç', merge_small_header2)
            ws.merge_range('A32:G32', u'Kullanım', merge_small_header2)
            ws.merge_range('A33:G33', u'Verinin Kökeni', merge_small_header2)
            
            ws.merge_range('A34:G34', u'Verinin Eksiksizliği (Completeness)', merge_small_header2)
            ws.merge_range('A35:G35', u'Fazlalık', xsmall_header_right)
            ws.merge_range('A36:G36', u'Eksiklik', xsmall_header_right)
            
            ws.merge_range('A37:G37', u'Mantıksal Tutarlılık (Logical Consistency)', merge_small_header2)
            ws.merge_range('A38:G38', u'Kavramsal Tutarlılık', xsmall_header_right)
            ws.merge_range('A39:G39', u'Tanım Kümesi Tutarlılığı', xsmall_header_right)
            ws.merge_range('A40:G40', u'Format Tutarlılığı', xsmall_header_right)
            ws.merge_range('A41:G41', u'Topoloji Tutarlılığı', xsmall_header_right)

            ws.merge_range('A42:G42', u'Konumsal Doğruluk (Positional Accuracy)', merge_small_header2)
            ws.merge_range('A43:G43', u'Mutlak Doğruluk', xsmall_header_right)
            ws.merge_range('A44:G44', u'Bağıl Doğruluk', xsmall_header_right)
            ws.merge_range('A45:G45', u'Raster Veri Konum Doğruluğu', xsmall_header_right)

            ws.merge_range('A46:G46', u'Zamansal Doğruluk (Temporal Accuracy)', merge_small_header2)
            ws.merge_range('A47:G47', u'İlgili Zamandaki Doğruluk', xsmall_header_right)
            ws.merge_range('A48:G48', u'Zamansal Tutarlılık', xsmall_header_right)
            ws.merge_range('A49:G49', u'Zamansal Geçerlilik', xsmall_header_right)

            ws.merge_range('A50:G50', u'Tematik Doğruluk (Thematic Accuracy)', merge_small_header2)
            ws.merge_range('A51:G51', u'Sınıflandırma Doğruluğu', xsmall_header_right)
            ws.merge_range('A52:G52', u'Nicel öznitelik bilgilerinin doğruluğu', xsmall_header_right)
            ws.merge_range('A53:G53', u'Nicel olmayan öznitelik bilgilerinin doğruluğu', xsmall_header_right)
            ws.merge_range('A54:G55', u'Açıklama', merge_small_header_left)


            if self.vk_amac is not None:
                h = 15*len(self.vk_amac)/100
                ws.merge_range('H31:U31', self.vk_amac.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(30, h)
            else:
                ws.merge_range('H31:U31', u'')
                
            #     ws.merge_range('H31:U31', self.vk_amac.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H31:U31', u'', f_data_emty)

            if self.vk_kullanim is not None:
            #     ws.merge_range('H32:U32', u'', f_data_right)
            #     ws.write('H32',self.vk_kullanim.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H32:U32', u'', f_data_emty)

                h = 15*len(self.vk_kullanim)/100
                ws.merge_range('H32:U32', self.vk_kullanim.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(31, h)
            else:
                ws.merge_range('H32:U32', u'')
            
            if self.vk_kokeni is not None:
                h = 15*len(self.vk_kokeni)/100
                ws.merge_range('H33:U33', self.vk_kokeni.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(32, h)
            else:
                ws.merge_range('H33:U33', u'')
                
                
            #     ws.merge_range('H33:U33', self.vk_kokeni.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H33:U33', u'', f_data_emty)

            if self.vk_copleteness_fazlalik is not None:
                h = 15*len(self.vk_copleteness_fazlalik)/100
                ws.merge_range('H34:U34', self.vk_copleteness_fazlalik.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(33, h)
            else:
                ws.merge_range('H34:U34', u'')
                
            #     ws.merge_range('H34:U34', self.vk_copleteness_fazlalik.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H34:U34', u'', f_data_emty)
            
            if self.vk_fazlalik_yeni is not None:
                ws.merge_range('H35:U35', self.vk_fazlalik_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H35:U35', u'', f_data_emty)
            
            if self.veri_eksiklik_additional is not None:
                ws.merge_range('H36:U36', self.veri_eksiklik_additional.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H36:U36', u'', f_data_emty)

            if self.vk_lc_kavramsal_tutarlilik is not None:
                ws.merge_range('H37:U37', self.vk_lc_kavramsal_tutarlilik.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H37:U37', u'', f_data_emty)

            if self.vk_kavramsal_yeni is not None:
                ws.merge_range('H38:U38', self.vk_kavramsal_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H38:U38', u'', f_data_emty)

            if self.vk_tanim_kumesi_yeni is not None:
                ws.merge_range('H39:U39', self.vk_tanim_kumesi_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H39:U39', u'', f_data_emty)

            if self.vk_format_tutarlilik_yeni is not None:
                ws.merge_range('H40:U40', self.vk_format_tutarlilik_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H40:U40', u'', f_data_emty)
            
            if self.vk_topoloji_tutarlilik_yeni is not None:
                ws.merge_range('H41:U41', self.vk_topoloji_tutarlilik_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H41:U41', u'', f_data_emty)

            if self.vk_pa_mutlak_dogruluk is not None:
                h = 15*len(self.vk_pa_mutlak_dogruluk)/100
                ws.merge_range('H42:U42', self.vk_pa_mutlak_dogruluk.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(41, h)
            else:
                ws.merge_range('H42:U42', u'')                
                
            #     ws.merge_range('H42:U42', self.vk_pa_mutlak_dogruluk.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H42:U42', u'', f_data_emty)
            
            if self.vk_konumsal_mutlak_dogruluk_yeni is not None:
                ws.merge_range('H43:U43', self.vk_konumsal_mutlak_dogruluk_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H43:U43', u'', f_data_emty)
            
            if self.vk_konumsal_bagil_dogruluk_yeni is not None:
                ws.merge_range('H44:U44', self.vk_konumsal_bagil_dogruluk_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H44:U44', u'', f_data_emty)
            
            if self.vk_konumsal_raster_veri_konum_yeni is not None:
                ws.merge_range('H45:U45', self.vk_konumsal_raster_veri_konum_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H45:U45', u'', f_data_emty)

            if self.vk_ta_ilgili_zamandaki_dogruluk is not None:
                h = 15*len(self.vk_ta_ilgili_zamandaki_dogruluk)/100
                ws.merge_range('H46:U46', self.vk_ta_ilgili_zamandaki_dogruluk.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(45, h)
            else:
                ws.merge_range('H46:U46', u'')
                
            #     ws.merge_range('H46:U46', self.vk_ta_ilgili_zamandaki_dogruluk.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H46:U46', u'', f_data_emty)

            if self.vk_zamansal_ilgili_yeni is not None:
                ws.merge_range('H47:U47', self.vk_zamansal_ilgili_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H47:U47', u'', f_data_emty)

            if self.vk_zamansal_tutarlilik_yeni is not None:
                ws.merge_range('H48:U48', self.vk_zamansal_tutarlilik_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H48:U48', u'', f_data_emty)
            
            if self.vk_zamansal_gecerlilik_yeni is not None:
                ws.merge_range('H49:U49', self.vk_zamansal_gecerlilik_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H49:U49', u'', f_data_emty)

            if self.vk_tema_siniflandirma_dogrulugu is not None:
                h = 15*len(self.vk_tema_siniflandirma_dogrulugu)/100
                ws.merge_range('H50:U50', self.vk_tema_siniflandirma_dogrulugu.decode('utf-8'), f_data_right)
                if h < 15:
                    h = 15
                else:
                    h = 30
                ws.set_row(49, h)
            else:
                ws.merge_range('H50:U50', u'')
                
            #     ws.merge_range('H50:U50', self.vk_tema_siniflandirma_dogrulugu.decode('utf-8'), f_data_right)
            # else:
            #     ws.merge_range('H50:U50', u'', f_data_emty)

            if self.vk_tematik_siniflandirma_yeni is not None:
                ws.merge_range('H51:U51', self.vk_tematik_siniflandirma_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H51:U51', u'', f_data_emty)

            if self.vk_tematik_nicel_yeni is not None:
                ws.merge_range('H52:U52', self.vk_tematik_nicel_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H52:U52', u'', f_data_emty)

            if self.vk_tematik_nicel_olmayan_yeni is not None:
                ws.merge_range('H53:U53', self.vk_tematik_nicel_olmayan_yeni.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H53:U53', u'', f_data_emty)
            
            if self.vk_aciklama is not None:
                ws.merge_range('H54:U55', self.vk_aciklama.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('H54:U55', u'', f_border)
            ws.merge_range('A56:U56', u'')

            # çalışma grubu
            ws.merge_range('A57:U57', u'Çalışma Grubu', merge_header_format2)
            ws.merge_range('A58:F58', u'Grup Adı', merge_small_header2)
            ws.merge_range('A59:F59', u'Adı Soyadı', merge_small_header2)
            ws.merge_range('A60:F60', u'Görevi', merge_small_header2)
            ws.merge_range('A61:F61', u'Tarih', merge_small_header2)

            ws.merge_range('G58:U58', u'(Analiz Grubu Adı)', f_comment)
            ws.merge_range('G59:U59', u'(Analizi Gerçekleştiren Proje Uzmanı)', f_comment)
            ws.merge_range('G60:U60', u'(Proje Uzmanı Görevi)', f_comment)
            ws.merge_range('G61:U61', u'(Analiz Tarihi)', f_comment)
            ws.merge_range('A62:U62', u'*Tespit çalışması ile ilgili olarak Coğrafi Veri Analiz Formu kapsamında elde edilen bilgiler kurum temsilcilerinin beyanına dayanmaktadır.', f_comment_left)


            if self.tucbs_tema_harici is False:
                ws.merge_range('N16:U16', u'', merge_small_header2)
                ws.merge_range('N17:U17', u'', merge_small_header2)
                ws.merge_range('G12:Q12', self.katman_adi.decode('utf-8'), f_data_left)
                if self.tucbs_veri_temasi is not None:
                    ws.merge_range('G10:U10', self.tucbs_veri_temasi.decode('utf-8'), f_data_right)
                else:
                    ws.merge_range('G10:U10', '', f_data_emty)

                #if self.katman_durumu:
                ws.write_rich_string('R12', merge_small_header2,u'Var(', f_red, 'X', merge_small_header2,')',merge_small_header)
                ws.write('S12', u'Yok( )',merge_small_header) 
                #else:
                    # ws.write('R12', u'Var( )',merge_small_header) 
                    # ws.write_rich_string('S12', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',merge_small_header)
                
                if self.tucbs_uygunluk:
                    ws.write_rich_string('T12', merge_small_header2,u'Uygun(', f_red, 'X', merge_small_header2,')',merge_small_header)
                    ws.write('U12', u'Uygun Değil( )',merge_small_header) 
                else:
                    ws.write('T12', u'Uygun( )',merge_small_header) 
                    ws.write_rich_string('U12', merge_small_header2,u'Uygun Değil(', f_red, 'X', merge_small_header2,')',merge_small_header)
                ws.write('G16', u'Var( )',merge_small_header) 
                ws.write_rich_string('K16', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',merge_small_header)
                if self.inspire_uygunluk:
                    ws.write_rich_string('N18', merge_small_header2,u'Evet (', f_red, u'X', merge_small_header2,u') Hayır( )',f_border_center)
                else: 
                    ws.write_rich_string('N18', merge_small_header2,u'Evet ( ) Hayır (', f_red, u'X', merge_small_header2,')',f_border_center)              
            else:
                ws.merge_range('G10:U10', u'', merge_small_header)
                ws.write('R12', u'Var( )',merge_small_header) 
                ws.write_rich_string('S12', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',merge_small_header) 
                ws.write('T12', u'Uygun( )',merge_small_header) 
                ws.write('U12', u'Uygun Değil( )',merge_small_header) 
                ws.merge_range('G12:Q12', u'', merge_small_header)

            if self.tucbs_tema_harici:
                #if self.katman_durumu:
                ws.write_rich_string('G16', merge_small_header2,u'Var(', f_red, 'X', merge_small_header2,')',merge_small_header)
                ws.write('K16', u'Yok( )',merge_small_header) 
                # else:
                #     ws.write_rich_string('K16', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',merge_small_header)
                #     ws.write('G16', u'Var( )',merge_small_header) 
                
                ws.merge_range('N16:U16', self.katman_adi.decode('utf-8'), f_data_right)
                
                if self.inspire_katmani is not None:
                    ws.merge_range('N17:U17', self.inspire_katmani.decode('utf-8'), f_data_right)
                else:
                    ws.merge_range('N17:U17', '', f_data_emty)
                
                if self.inspire_uygunluk:
                    ws.write_rich_string('N18', merge_small_header2,u'Evet (', f_red, u'X', merge_small_header2,u') Hayır( )',f_border_center)
                else: 
                    ws.write_rich_string('N18', merge_small_header2,u'Evet ( ) Hayır (', f_red, u'X', merge_small_header2,')',f_border_center)
            else:
                pass
                
            wb.close()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
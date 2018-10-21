# -*- coding: utf-8 -*-
import xlsxwriter
from pgget import Connection
cnn = Connection()

class CografiVeriFormu:
    def __init__(self, bakanlik, adi, birim, tucbs_katmani, katman_adi, katman_durumu, tucbs_uygunluk, veri_turu, veri_tipi, veri_adedi, veri_formati, projeksiyon,
                    datum, olcek_duzey, veri_guncelleme_periyod, son_veri_guncelleme_tarih, veri_envanteri_aciklama, tucbs_tema_harici, inspire_katmani, inspire_uygunluk,
                    katman_aciklama, tesim_alindi, teslim_formati, teslim_alinan_veri_sayisi):
        self.bakanlik = bakanlik
        self.adi = adi
        self.birim = birim

        # tucbs temasinin kod tablosundan çekilmesi
        if tucbs_katmani is not None:
            self.tucbs_veri_temasi = cnn.getsinglekoddata('kod_tucbs_tema', 'tema_adi', 'objectid='+str(tucbs_katmani))
        else:
            self.tucbs_veri_temasi = None
        self.katman_adi = katman_adi
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
            self.teslim_formati = cnn.getsinglekoddata('kod_ek_2_veri_turu', 'kod', 'objectid='+str(teslim_formati))
        else:
            self.teslim_formati = None
        
        
    
    def createExcelFile(self):
        wb = xlsxwriter.Workbook(r'created_excels\demo.xlsx')
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
        ws.set_column('S:S', 11.43)
        ws.set_column('T:T', 18.57)
        ws.set_column('U:U', 16.23)
        ws.set_row(5, 21)
        ws.set_row(4, 5)
        ws.set_row(14, 5)
        ws.set_row(18, 25)
        ws.set_row(17, 21)
        ws.set_row(20, 24.75)
        ws.set_row(21, 14.25)
        ws.set_row(22, 75)
        ws.set_row(24, 5)
        ws.set_row(25, 21)

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

        merge_small_header2 = wb.add_format()
        merge_small_header2.set_font_size(11)
        merge_small_header2.set_bold()
        merge_small_header2.set_border()

        f_data_right = wb.add_format()
        f_data_right.set_bold()
        f_data_right.set_font_color('red')
        f_data_right.set_align('right')
        f_data_right.set_border()
        f_data_right.set_text_wrap()

        f_data_left = wb.add_format()
        f_data_left.set_bold()
        f_data_left.set_font_color('red')
        f_data_left.set_align('left')
        f_data_left.set_border()
        f_data_left.set_text_wrap()

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

        f_comment = wb.add_format()
        f_comment.set_border(),f_border_center
        f_comment.set_color('gray')
        f_comment.set_italic()
        f_comment.set_font_size(9)
        f_comment.set_align('right')

        f_comment_left = wb.add_format()
        f_comment_left.set_border()
        f_comment_left.set_color('gray')
        f_comment_left.set_italic()
        f_comment_left.set_font_size(9)
        

        ws.insert_image('A1', r"logo\csb.jpg", {'x_offset': 20,'y_offset': 7,'x_scale': 1.6})
        ws.merge_range('A1:D4','',merge_header_format)

        ws.merge_range('E1:Q4',u'Coğrafi Veri Analiz Formu', merge_header_format)

        ws.merge_range('R1:S1', u'Revizyon Numarası', merge_small_header)
        ws.merge_range('R2:S2', '',merge_small_header)
        ws.merge_range('R3:S3', u'Revizyon Tarihi', merge_small_header)
        ws.merge_range('R4:S4', '',merge_small_header)
        ws.merge_range('T1:U4', '',merge_small_header)
        ws.merge_range('A5:S5', '',merge_small_header)
        ws.merge_range('A6:F6', u'Genel Bilgiler',merge_header_format2)
        ws.merge_range('G6:U6', u'TUCBS Analiz Portalı üzerinden doldurulacak',f_comment)
        ws.merge_range('G13:U14', '',merge_small_header)
        ws.merge_range('A15:S15', '',merge_small_header)


        ws.insert_image('T1', r"logo\tucbs2.jpg", {'x_offset': 10,'y_offset': 3,'x_scale': 0.5,'y_scale': 0.5})

        ws.merge_range('A7:F7', u'Bakanlık', merge_small_header2)
        ws.merge_range('A8:F8', u'Genel Müdürlük / Belediye', merge_small_header2)
        ws.merge_range('A9:F9', u'Birim', merge_small_header2)
        ws.merge_range('A10:F10', u'TUCBS Coğrafi Veri Teması', merge_small_header2)
        ws.merge_range('A11:F14', u'TUCBS Veri Katmanları', merge_small_header)
        ws.merge_range('A16:F18', u'Tema Harici Üretilen Veri Katmanı Var mı?', merge_small_header)
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
        ws.merge_range('J21:K22', u'',f_border_center)
        ws.merge_range('L21:P22', u'',f_border_center)
        ws.merge_range('Q21:R22', u'',f_border_center)
        ws.merge_range('S21:S22', u'',f_border_center)
        ws.merge_range('T21:T22', u'',f_border_center)
        ws.merge_range('U21:U22', u'',f_border_center)
        ws.write('A21',u'Katman Adı', f_border_center)
        ws.write_rich_string('D21', merge_small_header, u'Veri Türü', f_comment, u' (Dijital Veri / Basılı Veri)', f_border_center)
        ws.write_rich_string('G21', merge_small_header, u'Veri Tipi', f_comment, u' (Coğrafi Veri / Sözel Veri)', f_border_center)
        ws.write('J21', u'Veri Adetleri', merge_small_header)
        ws.write('L21', u'Veri Formatı', merge_small_header)
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
            ws.merge_range('J23:K23',self.veri_adedi.decode('utf-8'), f_data_center)
        else:
            ws.merge_range('J23:K23', u'', f_data_emty)
        # veri formati
        if self.veri_formati is not None:
            ws.merge_range('L23:P23',self.veri_formati.decode('utf-8'), f_data_center)
        else:
            ws.merge_range('L23:P23', u'', f_data_emty)
        # projeksyion datum
        if self.projeksiyon is not None or self.datum is not None:
            ws.merge_range('Q23:R23',self.projeksiyon.decode('utf-8') + " " + self.datum.decode('utf-8'), f_data_center)
        else:
            ws.merge_range('Q23:R23', u'', f_data_emty)
        # ölçek düzey çözünürlük
        if self.olcek_duzey is not None:
            ws.write('S23',self.olcek_duzey.decode('utf-8'), f_data_center)
        else:
            ws.write('S23', u'', f_data_emty)
        # veri güncelleme
        if self.veri_guncelleme_periyod is not None:
            ws.write('T23',self.veri_guncelleme_periyod.decode('utf-8'), f_data_center)
        else:
            ws.write('T23', u'', f_data_emty)
        # son veri güncelleme tarihi
        if self.son_veri_guncelleme_tarih is not None:
            ws.write('U23',self.son_veri_guncelleme_tarih.strftime('%Y-%m-%d'), f_data_center)
        else:
            ws.write('U23', u'', f_data_emty)

        # katman aciklama
        if self.veri_envanteri_aciklama is not None:
            ws.set_row(23, 25)
            ws.merge_range('A24:U24', self.veri_envanteri_aciklama.decode('utf-8'), f_data_left)
        else:
            ws.merge_range('A24:U24', u'', merge_small_header2)
            ws.set_row(23, 5)

        # veri teslimi
        ws.merge_range('A26:U26', u'', merge_small_header)
        ws.write_rich_string('A26', merge_header_format2, u'Veri Teslimi', f_comment_left, u'  (Teslim alınan veriler için doldurulacaktır)',f_border)
        ws.merge_range('A27:F27', u'Veri Katmanı', merge_small_header)
        ws.merge_range('G27:L27', u'', merge_small_header)
        ws.write_rich_string('G27', merge_small_header, u'Veri Tipi', f_comment, u' (Coğrafi Veri / Sözel Veri)', f_border_center)
        ws.merge_range('M27:S27', u'Veri Formatı', merge_small_header)
        ws.merge_range('T27:U27', u'Veri Sayısı', merge_small_header)
        ws.merge_range('A28:F28', u'', merge_small_header)
        ws.merge_range('G28:L28', u'', merge_small_header)
        ws.merge_range('M28:S28', u'', merge_small_header)
        ws.merge_range('T28:U28', u'', merge_small_header)


        if self.teslim_alindi:
            ws.write('A28', self.katman_adi.decode('utf-8'),f_data_center)
            if self.veri_tipi:
                ws.write('G28', self.veri_tipi.decode('utf-8'),f_data_center)
            else:
                ws.write('G28', u'', f_data_emty)

            pass
        else:
            pass

        ws.merge_range('G16:J16', u'Var( )', merge_small_header2)
        ws.merge_range('K16:M16', u'Yok( )', merge_small_header2)
        ws.merge_range('N18:U18', u'Evet ( )    Hayır( )', merge_small_header)


        if self.tucbs_tema_harici is False:
            ws.merge_range('N16:U16', u'', merge_small_header2)
            ws.merge_range('N17:U17', u'', merge_small_header2)
            ws.merge_range('G12:Q12', self.katman_adi.decode('utf-8'), f_data_left)
            if self.tucbs_veri_temasi is not None:
                ws.merge_range('G10:U10', self.tucbs_veri_temasi.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('G10:U10', '', f_data_emty)

            if self.katman_durumu:
                ws.write_rich_string('R12', merge_small_header2,u'Var(', f_red, 'X', merge_small_header2,')',f_border)
                ws.write('S12', u'Yok( )',merge_small_header2) 
            else:
                ws.write('R12', u'Var( )',merge_small_header2) 
                ws.write_rich_string('S12', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',f_border)
            
            if self.tucbs_uygunluk:
                ws.write_rich_string('T12', merge_small_header2,u'Uygun(', f_red, 'X', merge_small_header2,')',f_border)
                ws.write('U12', u'Uygun Değil( )',merge_small_header2) 
            else:
                ws.write('T12', u'Uygun( )',merge_small_header2) 
                ws.write_rich_string('U12', merge_small_header2,u'Uygun Değil(', f_red, 'X', merge_small_header2,')',f_border)
        else:
            ws.merge_range('G10:U10', u'', merge_small_header2)
            ws.write('R12', u'Var( )',merge_small_header2) 
            ws.write('S12', u'Yok( )',merge_small_header2) 
            ws.write('T12', u'Uygun( )',merge_small_header2) 
            ws.write('U12', u'Uygun Değil( )',merge_small_header2) 
            ws.merge_range('G12:Q12', u'', merge_small_header2)

        if self.tucbs_tema_harici:
            if self.katman_durumu:
                ws.write_rich_string('G16', merge_small_header2,u'Var(', f_red, 'X', merge_small_header2,')',f_border)
                ws.write('K16', u'Yok( )',merge_small_header2) 
            else:
                ws.write_rich_string('K16', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',f_border)
                ws.write('G16', u'Var( )',merge_small_header2) 
            
            ws.merge_range('N16:U16', self.katman_adi.decode('utf-8'), f_data_right)
            
            if self.inspire_katmani is not None:
                ws.merge_range('N17:U17', self.inspire_katmani.decode('utf-8'), f_data_right)
            else:
                ws.merge_range('N17:U17', '', f_data_emty)
            
            if self.inspire_uygunluk:
                ws.write_rich_string('N18', merge_small_header2,u'Evet (', f_red, 'X', merge_small_header2,')    Hayır( )',f_border_center)
            else: 
                ws.write_rich_string('N18', merge_small_header2,u'Evet ( )     Hayır (', f_red, 'X', merge_small_header2,')',f_border_center)
        else:
            pass
            
        wb.close()
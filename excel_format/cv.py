# -*- coding: utf-8 -*-
import xlsxwriter
from pgget import Connection
cnn = Connection()

class CografiVeriFormu:
    def __init__(self, bakanlik, adi, birim, tucbs_katmani, katman_adi, katman_durumu, tucbs_uygunluk, veri_turu, veri_tipi, veri_adedi, veri_formati, projeksiyon,
                    datum, olcek_duzey, veri_guncelleme_periyod, son_veri_guncelleme_tarih, veri_envanteri_aciklama, tucbs_tema_harici, inspire_katmani, inspire_uygunluk):
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

        merge_small_header2 = wb.add_format()
        merge_small_header2.set_font_size(11)
        merge_small_header2.set_bold()
        merge_small_header2.set_border()

        f_data_right = wb.add_format()
        f_data_right.set_bold()
        f_data_right.set_font_color('red')
        f_data_right.set_align('right')
        f_data_right.set_border()

        f_data_left = wb.add_format()
        f_data_left.set_bold()
        f_data_left.set_font_color('red')
        f_data_left.set_align('left')
        f_data_left.set_border()

        f_red = wb.add_format()
        f_red.set_font_color('red')
        f_red.set_bold()
        f_red.set_border()

        f_border = wb.add_format()
        f_border.set_border()

        f_border_center = wb.add_format()
        f_border_center.set_border()
        f_border_center.set_align('center')



        f_data_emty = wb.add_format()
        f_data_emty.set_bg_color('yellow')
        f_data_emty.set_border()

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
        ws.merge_range('G6:U6', '',merge_small_header)
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
        if self.tucbs_veri_temasi is not None:
            ws.merge_range('G10:U10', self.tucbs_veri_temasi.decode('utf-8'), f_data_right)
        else:
            ws.merge_range('G10:U10', '', f_data_emty)

        ws.merge_range('G11:Q11', u'Veri Katman Adı', merge_small_header2)
        ws.merge_range('R11:S11', u'Veri Katman Durumu', merge_small_header)
        ws.merge_range('T11:U11', u'TUCBS Standartlarına Uygunluk', merge_small_header)
        ws.merge_range('G17:M17', u'Katmanın INSPIRE\' a Uygun Tema Adı', merge_small_header2)
        ws.merge_range('G18:M18', u'Katman INSPIRE Standartlarına Uygun Mu?', merge_small_header2)


        ws.merge_range('G12:Q12', self.katman_adi.decode('utf-8'), f_data_left)

        if self.katman_durumu is True:
            ws.write_rich_string('R12', merge_small_header2,u'Var(', f_red, 'X', merge_small_header2,')',f_border)
            ws.write('S12', u'Yok( )',merge_small_header2) 
        else:
            ws.write('R12', u'Var( )',merge_small_header2) 
            ws.write_rich_string('S12', merge_small_header2,u'Yok(', f_red, 'X', merge_small_header2,')',f_border)
        
        if self.tucbs_uygunluk is True:
            ws.write_rich_string('T12', merge_small_header2,u'Uygun(', f_red, 'X', merge_small_header2,')',f_border)
            ws.write('U12', u'Uygun Değil( )',merge_small_header2) 
        else:
            ws.write('T12', u'Uygun( )',merge_small_header2) 
            ws.write_rich_string('U12', merge_small_header2,u'Uygun Değil(', f_red, 'X', merge_small_header2,')',f_border)

        ws.merge_range('G16:J16', u'Var( )', merge_small_header2)
        ws.merge_range('K16:M16', u'Yok( )', merge_small_header2)
        ws.merge_range('N18:U18', u'Evet ( )    Hayır( )', merge_small_header2)
        ws.merge_range('N16:U16', u'', merge_small_header2)
        ws.merge_range('N17:U17', u'', merge_small_header2)


        if self.tucbs_tema_harici is True:
            if self.katman_durumu is True:
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
            
            if self.inspire_uygunluk is True:
                ws.write_rich_string('N18', merge_small_header2,u'Evet (', f_red, 'X', merge_small_header2,')    Hayır( )',f_border_center)
            else: 
                ws.write_rich_string('N18', merge_small_header2,u'Evet ( )     Hayır (', f_red, 'X', merge_small_header2,')',f_border_center)
        else:
            pass
            
        wb.close()
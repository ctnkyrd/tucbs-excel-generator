# -*- coding: utf-8 -*-
import xlsxwriter

class CografiVeriFormu:
    def __init__(self, bakanlik, adi, birim):
        self.bakanlik = bakanlik
        self.adi = adi
        self.birim = birim
    
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
        ws.insert_image('T1', r"logo\tucbs2.jpg", {'x_offset': 10,'y_offset': 3,'x_scale': 0.5,'y_scale': 0.5})

        ws.merge_range('A7:F7', u'Bakanlık', merge_small_header2)
        ws.merge_range('A8:F8', u'Genel Müdürlük / Belediye', merge_small_header2)
        ws.merge_range('A9:F9', u'Birim', merge_small_header2)
        ws.merge_range('A10:F10', u'TUCBS Coğrafi Veri Teması', merge_small_header2)
        ws.merge_range('A11:F14', u'TUCBS Veri Katmanları', merge_small_header)
        # fill form
        ws.merge_range('G7:U7', self.bakanlik, f_data_right)
        ws.merge_range('G8:U8', self.adi, f_data_right)
        ws.merge_range('G9:U9', self.birim, f_data_right)

        wb.close()
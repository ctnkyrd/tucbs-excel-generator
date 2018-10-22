# -*- coding: utf-8 -*-
import xlsxwriter
from pgget import Connection
cnn = Connection()

class MetaveriFormu:
    def __init__(self, katman_adi, mv_metaveri_var, mv_standart, mv_yayinlaniyor, mv_cbs_gm_paylasim_var, metaveri_aciklama):
        
        
        self.katman_adi = katman_adi
        self.metaveri_aciklama = metaveri_aciklama

        if mv_metaveri_var is not None:
            self.mv_metaveri_var = True
        else:
            self.mv_metaveri_var = False

        if mv_standart is not None:
            self.mv_standart = cnn.getsinglekoddata('kod_ek_2_meta_veri_standart', 'kod', 'objectid='+str(mv_standart))
        else:
            self.mv_standart = None

        if mv_yayinlaniyor is not None:
            self.mv_yayinlaniyor = True
        else:
            self.mv_yayinlaniyor = False

        if mv_cbs_gm_paylasim_var is not None:
            self.mv_cbs_gm_paylasim_var =True
        else:
            self.mv_cbs_gm_paylasim_var = False

    def createExcelFile(self):
        try:

            # Dosya Olusturma
            workbook = xlsxwriter.Workbook('mv.xlsx')
            worksheet = workbook.add_worksheet()

            # Satir / Sutun Ayarlari

            worksheet.set_column('D:D', 11)
            worksheet.set_column('L:L', 7.57)
            worksheet.set_row(6, 3.75)
            worksheet.set_row(12, 33)

            # Formatlar
            merge_format = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})

            main_header_format = workbook.add_format({
                'font_size': 16,
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})

            header_format = workbook.add_format({
                'font_size': 16,
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'})

            text_format = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'})
            text_format.set_text_wrap()

            data_format_r = workbook.add_format({
                'font_color': 'red',
                'bold': 1,
                'border': 1,
                'align': 'right',
                'valign': 'vcenter'})
            data_format_r.set_text_wrap()

            data_format_c = workbook.add_format({
                'font_color': 'red',
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})
            data_format_c.set_text_wrap()

            data_empty = workbook.add_format({
                'bg_color': '#C5C5C5',
                'border': 1})

            # Baslik
            worksheet.insert_image('A1', r"logo\csb.jpg", {'x_offset': 70,'y_offset': 5,'x_scale': 1.25})
            worksheet.merge_range('A1:D4', '', merge_format)
            worksheet.merge_range('E1:J4', u'Metaveri Analiz Formu', main_header_format)
            worksheet.merge_range('K1:M1', u'Revizyon Numarası', merge_format)
            worksheet.merge_range('K2:M2', '', merge_format)
            worksheet.merge_range('K3:M3', u'Revizyon Tarihi', merge_format)
            worksheet.merge_range('K4:M4', '', merge_format)
            worksheet.insert_image('N1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})
            worksheet.merge_range('N1:P4', '', merge_format)
            worksheet.merge_range('A5:P5', '')

            # Metaveri Analizi
            worksheet.merge_range('A6:P6', 'Metaveri Analizi', header_format)
            worksheet.merge_range('A7:P7', '')

            worksheet.merge_range('A8:D8', u'Metaveri Analizine Konu Veri Katmanı', text_format)
            worksheet.merge_range('E8:P8', self.katman_adi.decode('utf-8'), data_format_r)

            worksheet.merge_range('A9:D9', u'Metaveri Var Mı?', text_format)
            if self.mv_metaveri_var is True:
                worksheet.merge_range('E9:P9', u'Evet', data_format_r)
            elif self.mv_metaveri_var is False:
                worksheet.merge_range('E9:P9', u'Hayır', data_format_r)
            else:
                worksheet.merge_range('E9:P9', u'', data_empty)

            worksheet.merge_range('A10:D10', u'Metaveri Hangi Standarta Uygun Üretiliyor?', text_format)
            if self.mv_standart is True:
                worksheet.merge_range('E9:P9', u'Evet', data_format_r)
            elif self.mv_standart is False:
                worksheet.merge_range('E9:P9', u'Hayır', data_format_r)
            else:
                worksheet.merge_range('E9:P9', u'', data_empty)

            worksheet.merge_range('A11:D11', u'Metaveri Yayınlanıyor Mu?', text_format)
            if self.mv_yayinlaniyor is True:
                worksheet.merge_range('E9:P9', u'Evet', data_format_r)
            elif self.mv_yayinlaniyor is False:
                worksheet.merge_range('E9:P9', u'Hayır', data_format_r)
            else:
                worksheet.merge_range('E9:P9', u'', data_empty)

            worksheet.merge_range('A12:D12', u'CBS Genel Müdürlüğü ile Paylaşımı Var Mı? ', text_format)
            if self.mv_cbs_gm_paylasim_var is True:
                worksheet.merge_range('E9:P9', u'Evet', data_format_r)
            elif self.mv_cbs_gm_paylasim_var is False:
                worksheet.merge_range('E9:P9', u'Hayır', data_format_r)
            else:
                worksheet.merge_range('E9:P9', u'', data_empty)

            worksheet.merge_range('A13:D13', u'Açıklama', text_format)
            worksheet.merge_range('E13:P13', self.metaveri_aciklama.decode('utf-8'), data_format_r)

            # Calisma Grubu
            worksheet.merge_range('A14:P14', u'Çalışma Grubu', header_format)
            worksheet.merge_range('A15:P15', '')
            worksheet.merge_range('A16:D16', u'Grup Adı', text_format)
            worksheet.merge_range('E16:P16', u'', data_format_r)
            worksheet.merge_range('A17:D17', u'Adı Soyadı', text_format)
            worksheet.merge_range('E17:P17', u'', data_format_r)
            worksheet.merge_range('A18:D18', u'Görevi', text_format)
            worksheet.merge_range('E18:P18', u'', data_format_r)
            worksheet.merge_range('A19:D19', u'Tarih', text_format)
            worksheet.merge_range('E19:P19', u'', data_format_r)

            workbook.close()
        except BaseException as ex:
            print ex
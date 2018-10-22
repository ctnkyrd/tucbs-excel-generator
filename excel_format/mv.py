# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class MetaveriFormu:
    def __init__(self, katman_adi, mv_metaveri_var, mv_standart, mv_yayinlaniyor, mv_cbs_gm_paylasim_var, metaveri_aciklama,
                adi, tucbs_katmani, inspire_katmani):
        
        
        self.katman_adi = katman_adi.rstrip()
        self.metaveri_aciklama = metaveri_aciklama
        self.mv_metaveri_var = mv_metaveri_var

        if mv_standart is not None:
            self.mv_standart = cnn.getsinglekoddata('kod_ek_2_meta_veri_standart', 'kod', 'objectid='+str(mv_standart))
        else:
            self.mv_standart = None

        self.mv_yayinlaniyor = mv_yayinlaniyor
        self.mv_cbs_gm_paylasim_var = mv_cbs_gm_paylasim_var
        
        self.adi = adi.rstrip()
        if inspire_katmani is not None:
            self.inspire_katmani = cnn.getsinglekoddata('kod_inspire_tema', 'tema_adi', 'objectid='+str(inspire_katmani))
        else:
            self.inspire_katmani = None
        if tucbs_katmani is not None:
            self.tucbs_veri_temasi = cnn.getsinglekoddata('kod_tucbs_tema', 'tema_adi', 'objectid='+str(tucbs_katmani))
        else:
            self.tucbs_veri_temasi = None

    def createExcelFile(self):
        try:

            excelPath = "created_excels"+"\\"+self.adi+"\\"+u"CV-SP-MV"
            excelName = u"TUCBS-MVAF-MetaveriAnalizFormu.xlsx"
            temaName = u"Tema Yok"
            katmanName = u"Katman Yok"
            if self.tucbs_veri_temasi is not None:
                temaName = self.tucbs_veri_temasi.decode('utf-8')
            elif self.inspire_katmani is not None:
                temaName = self.inspire_katmani.decode('utf-8')
            
            if self.katman_adi is not None:
                katmanName = self.katman_adi.decode('utf-8')
                if '/' in katmanName:
                    katmanName = katmanName.replace('/', '_')
            else:
                katmanName = u"Katman Yok"

            fullFolderPath = excelPath+"\\"+temaName.rstrip()+"\\"+katmanName
            if os.path.isdir(unicode(fullFolderPath)) is False:
                try:
                    os.makedirs(unicode(fullFolderPath))
                except BaseException as ex:
                    print fullFolderPath
                    print ex

            # Dosya Olusturma
            workbook = xlsxwriter.Workbook(fullFolderPath+"\\"+excelName)
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

            comment_format = workbook.add_format({
                'font_size': 9,
                'font_color': 'gray',
                'italic': True,
                'border': 1,
                'align': 'right',
                'valign': 'vcenter'
            })

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
            if self.katman_adi is not None:
                worksheet.merge_range('E8:P8', self.katman_adi.decode('utf-8'), data_format_r)
            else:
                worksheet.merge_range('E8:P8', u'', data_empty)

            worksheet.merge_range('A9:D9', u'Metaveri Var Mı?', text_format)
            if self.mv_metaveri_var:
                worksheet.merge_range('E9:P9', u'Evet', data_format_r)
            else:
                worksheet.merge_range('E9:P9', u'Hayır', data_format_r)            

            worksheet.merge_range('A10:D10', u'Metaveri Hangi Standarta Uygun Üretiliyor?', text_format)
            if self.mv_standart:
                worksheet.merge_range('E10:P10', u'Evet', data_format_r)
            else:
                worksheet.merge_range('E10:P10', u'Hayır', data_format_r)

            worksheet.merge_range('A11:D11', u'Metaveri Yayınlanıyor Mu?', text_format)
            if self.mv_yayinlaniyor:
                worksheet.merge_range('E11:P11', u'Evet', data_format_r)
            else:
                worksheet.merge_range('E11:P11', u'Hayır', data_format_r)

            worksheet.merge_range('A12:D12', u'CBS Genel Müdürlüğü ile Paylaşımı Var Mı? ', text_format)
            if self.mv_cbs_gm_paylasim_var:
                worksheet.merge_range('E12:P12', u'Evet', data_format_r)
            else:
                worksheet.merge_range('E12:P12', u'Hayır', data_format_r)

            worksheet.merge_range('A13:D13', u'Açıklama', text_format)
            if self.metaveri_aciklama is not None:
                worksheet.merge_range('E13:P13', self.metaveri_aciklama.decode('utf-8'), data_format_r)
            else:
                worksheet.merge_range('E13:P13', u'', data_format_r)
            worksheet.merge_range('A14:P14', '')

            # Calisma Grubu
            worksheet.merge_range('A15:P15', u'Çalışma Grubu', header_format)            
            worksheet.merge_range('A16:D16', u'Grup Adı', text_format)
            worksheet.merge_range('E16:P16', u'(Analiz Grubu Adı)', comment_format)
            worksheet.merge_range('A17:D17', u'Adı Soyadı', text_format)
            worksheet.merge_range('E17:P17', u'(Analizi Gerçekleştiren Proje Uzmanı)', comment_format)
            worksheet.merge_range('A18:D18', u'Görevi', text_format)
            worksheet.merge_range('E18:P18', u'(Proje Uzmanı Görevi)', comment_format)
            worksheet.merge_range('A19:D19', u'Tarih', text_format)
            worksheet.merge_range('E19:P19', u'(Analiz Tarihi)', comment_format)

            workbook.close()
        except BaseException as ex:
            print ex
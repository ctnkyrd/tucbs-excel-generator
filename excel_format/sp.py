# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class ServisPaylasimFormu:
    def __init__(self, katman_adi, servis_var, servis_ogc_uyumlu, servis_atlas_uyumlu, servis_wms_var, servis_wfs_var, servis_wms_version,
                servis_wfs_version, servis_aciklama, servis_yayin_platformu, sp_olmamasi_personel,
                sp_olmamasi_mevzuat, sp_olmamasi_donanim, sp_olmamasi_diger, sp_olmamasi_aciklama, adi, tucbs_katmani, inspire_katmani):
        
        self.servis_turu = ''
        self.katman_adi = katman_adi.rstrip()
        self.servis_var = servis_var
        self.servis_ogc_uyumlu = servis_ogc_uyumlu
        self.servis_atlas_uyumlu = servis_atlas_uyumlu
        self.servis_wms_var = servis_wms_var
        self.servis_wfs_var = servis_wfs_var
        self.servis_wms_version = servis_wms_version
        self.servis_wfs_version = servis_wfs_version
        self.servis_aciklama = servis_aciklama
        self.servis_yayin_platformu = servis_yayin_platformu
        self.sp_olmamasi_personel = sp_olmamasi_personel
        self.sp_olmamasi_mevzuat = sp_olmamasi_mevzuat
        self.sp_olmamasi_donanim = sp_olmamasi_donanim
        self.sp_olmamasi_diger = sp_olmamasi_diger
        self.sp_olmamasi_aciklama = sp_olmamasi_aciklama
        self.adi = adi.rstrip()
        if inspire_katmani is not None:
            self.inspire_katmani = cnn.getsinglekoddata('kod_inspire_tema', 'tema_adi', 'objectid='+str(inspire_katmani))
        else:
            self.inspire_katmani = None
        if tucbs_katmani is not None:
            self.tucbs_veri_temasi = cnn.getsinglekoddata('kod_tucbs_tema', 'tema_adi', 'objectid='+str(tucbs_katmani))
        else:
            self.tucbs_veri_temasi = None
        if self.servis_wms_var and self.servis_wfs_var is False:
            self.servis_turu = self.servis_turu + 'WMS'
        if self.servis_wms_var is False and self.servis_wfs_var:
            self.servis_turu = self.servis_turu + 'WFS'
        elif self.servis_wms_var and self.servis_wfs_var:
            self.servis_turu = self.servis_turu + 'WMS + WFS'
        else:
            self.servis_turu = ''


    def createExcelFile(self):
        try:

            excelPath = "created_excels"+"\\"+self.adi+"\\"+u"CV-SP-MV"
            excelName = u"TUCBS-VSPAF-VeriServislerivePaylaşımıAnalizFormu.xlsx"
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

            worksheet.set_column('A:A', 41.43)
            worksheet.set_column('B:B', 13.43)
            worksheet.set_column('C:C', 11.29)
            worksheet.set_column('D:D', 10.43)
            worksheet.set_column('E:E', 14.29)
            worksheet.set_column('F:F', 8.14)
            worksheet.set_column('G:G', 15)
            worksheet.set_column('H:H', 13.86)
            worksheet.set_column('I:I', 10.29)
            worksheet.set_column('J:J', 5.57)
            worksheet.set_column('K:K', 9.43)
            worksheet.set_row(7, 39.75)
            worksheet.set_row(16, 41.25)
            worksheet.set_row(17, 41.25)
            worksheet.set_row(18, 41.25)
            worksheet.set_row(19, 21)
            worksheet.set_row(21, 28.5)
            worksheet.set_row(22, 28.5)
            worksheet.set_row(23, 26.25)
            worksheet.set_row(25, 24)

            # Formatlar
            merge_format = workbook.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})
            merge_format.set_text_wrap()

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

            comment_format_r = workbook.add_format({
                'font_size': 9,
                'font_color': 'gray',
                'italic': True,
                'border': 1,
                'align': 'right',
                'valign': 'vcenter'
            })

            comment_format_c = workbook.add_format({
                'font_size': 9,
                'font_color': 'gray',
                'italic': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            border_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Baslik
            worksheet.insert_image('A1', r"logo\csb.jpg", {'x_offset': 70,'y_offset': 5,'x_scale': 1.25})
            worksheet.merge_range('A1:A4', '', merge_format)
            worksheet.merge_range('B1:F4', u'', merge_format)
            worksheet.write_rich_string('B1', main_header_format, u'Coğrafi Veri Servisleri ve Paylaşımı Analiz Formu', comment_format_c, u'\n(Her Tema İçin Ayrı Ayrı Doldurulacaktır.)', border_format)
            worksheet.merge_range('G1:H1', u'Revizyon Numarası', merge_format)
            worksheet.merge_range('G2:H2', '', merge_format)
            worksheet.merge_range('G3:H3', u'Revizyon Tarihi', merge_format)
            worksheet.merge_range('G4:H4', '', merge_format)
            worksheet.insert_image('I1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})
            worksheet.merge_range('I1:K4', '', merge_format)
            worksheet.merge_range('A5:K5', '')

            # Veri Servisleri ve Paylaşımı Analizi
            worksheet.merge_range('A6:K6', u'Coğrafi Veri Servisleri ve Paylaşımı Analizi', header_format)
            worksheet.merge_range('A7:K7', u'Teknik Altyapı', text_format)
            worksheet.write_rich_string('A8', merge_format, u'Veri Katmanı',  comment_format_c, u'\n(Analiz Edilen Veri Katmanı Adı)', border_format)
            worksheet.write_rich_string('B8', merge_format, u'Servis Durumu', comment_format_c, u'\n(Var/Yok)', border_format)
            worksheet.write_rich_string('C8', merge_format, u'Servis Türü', comment_format_c, u'\n(WMS,WFS, vd.)', border_format)
            worksheet.write('D8', u'Servis Versiyonu', merge_format)
            worksheet.merge_range('E8:F8', '', merge_format)
            worksheet.write_rich_string('E8', merge_format, u'OGC Uyumluluğu',  comment_format_c, u'\n(OGC Test linki eklenecek)', comment_format_c, '\n(Uyumlu/Uyumsuz)', border_format)
            worksheet.merge_range('G8:H8', '', merge_format)
            worksheet.write_rich_string('G8', merge_format, u'ATLAS Uyumluluğu', comment_format_c, u'\n(Uyumlu/Uyumsuz)', border_format)
            worksheet.merge_range('I8:K8', '', merge_format)
            worksheet.write_rich_string('I8', merge_format, u'Bağlantı Linki', comment_format_c, u'\n(Uyumluluk Testleri İçin Servis Erişim Linkleri)', border_format)

            if self.katman_adi is not None:
                worksheet.write('A9', self.katman_adi.decode('utf-8'), data_format_c)
            else:
                worksheet.write('A9', u'', data_empty)
            
            if self.servis_var:
                worksheet.write('B9', u'Var', data_format_c)
            else:
                worksheet.write('B9', u'Yok', data_format_c)
            
            worksheet.write('C9', self.servis_turu, data_format_c)

            



            






            workbook.close()
        except BaseException as ex:
            print ex
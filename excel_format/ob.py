# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class OrganizasyonBirimiFormu:
    def __init__(self,bakanlik, adi, k_adi, fileCount, cbs_birimi_var, cbs_yeni_birim_gerekli, veri_uretim_var, tasra_teskilati_var, sadece_sistem_idamesi_var,
                    ihtiyaclar_sizin_onayinizdan_geciyor, personel_yeterli, personel_yetersizlik_kriterleri, personel_yetersizlik_oneri, 
                    cbs_farkindaligi_vari, sorunlar, cbs_birim_adi, yapilanma_olcegi, kurum_semasindaki_yeri, tasra_teskilat_yapilanmasi, cbs_yeni_birim_gorusler
                ):
        # from kurum clas
        self.bakanlik = bakanlik
        self.adi = adi
        self.k_adi = k_adi
        self.fileCount = fileCount

        # from table
        self.cbs_birimi_var = cbs_birimi_var
        self.cbs_yeni_birim_gerekli = cbs_yeni_birim_gerekli
        self.veri_uretim_var = veri_uretim_var
        self.tasra_teskilati_var = tasra_teskilati_var
        self.sadece_sistem_idamesi_var = sadece_sistem_idamesi_var
        self.ihtiyaclar_sizin_onayinizdan_geciyor = ihtiyaclar_sizin_onayinizdan_geciyor
        self.personel_yeterli = personel_yeterli
        self.personel_yetersizlik_kriterleri = personel_yetersizlik_kriterleri
        self.personel_yetersizlik_oneri = personel_yetersizlik_oneri
        self.cbs_farkindaligi_vari = cbs_farkindaligi_vari
        self.sorunlar = sorunlar
        self.cbs_birim_adi = cbs_birim_adi
        self.yapilanma_olcegi = yapilanma_olcegi
        self.kurum_semasindaki_yeri = kurum_semasindaki_yeri
        self.tasra_teskilat_yapilanmasi = tasra_teskilat_yapilanmasi
        self.cbs_yeni_birim_gorusler = cbs_yeni_birim_gorusler

        
    def createExcelFile(self):
        try:

            excelPath = "created_excels"+"\\"+self.k_adi.decode('utf-8')
            if self.fileCount == 0:
                excelName = u"OBAF.xlsx"
            else:
                excelName = u"OBAF_"+str(self.fileCount)+".xlsx"

            fullFolderPath = excelPath

            if os.path.isdir(unicode(fullFolderPath)) is False:
                try:
                    os.makedirs(unicode(fullFolderPath))
                except BaseException as ex:
                    print fullFolderPath
                    print ex
            
            wb = xlsxwriter.Workbook(fullFolderPath+"\\"+excelName)
            # Dosya Olusturma
            ws = wb.add_worksheet()

            # Satir / Sutun Ayarlari



            # Formatlar
            merge_format = wb.add_format({
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})

            main_header_format = wb.add_format({
                'font_size': 16,
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})

            header_format = wb.add_format({
                'font_size': 16,
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'})

            text_format = wb.add_format({
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'})
            text_format.set_text_wrap()

            data_format_r = wb.add_format({
                'font_color': 'red',
                'bold': 1,
                'border': 1,
                'align': 'right',
                'valign': 'vcenter'})
            data_format_r.set_text_wrap()

            data_format_c = wb.add_format({
                'font_color': 'red',
                'bold': 1,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'})
            data_format_c.set_text_wrap()

            data_empty = wb.add_format({
                'bg_color': '#C5C5C5',
                'border': 1})

            comment_format = wb.add_format({
                'font_size': 9,
                'font_color': 'gray',
                'italic': True,
                'border': 1,
                'align': 'right',
                'valign': 'vcenter'
            })

            ws.set_column('A:A', 38.29)
            ws.set_column('B:B', 13.71)
            ws.set_column('C:C', 14.00)
            ws.set_column('D:D', 15.86)
            ws.set_column('E:E', 17.29)
            ws.set_column('F:F', 17.57)
            
            ws.merge_range('A1:A4', u'', main_header_format)

            # Baslik
            ws.insert_image('A1', r"logo\csb.jpg", {'x_offset': 70,'y_offset': 5,'x_scale': 1.25})
            # worksheet.merge_range('A1:D4', '', merge_format)
            # worksheet.merge_range('E1:J4', u'CBS Organizasyon Birimleri ve İnsan Kaynakları Analiz Formu', main_header_format)
            # worksheet.merge_range('K1:M1', u'Revizyon Numarası', merge_format)
            # worksheet.merge_range('K2:M2', '', merge_format)
            # worksheet.merge_range('K3:M3', u'Revizyon Tarihi', merge_format)
            # worksheet.merge_range('K4:M4', '', merge_format)
            # worksheet.insert_image('N1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})
            # worksheet.merge_range('N1:P4', '', merge_format)
            # worksheet.merge_range('A5:P5', '')

            

            wb.close()
        except BaseException as ex:
            print ex
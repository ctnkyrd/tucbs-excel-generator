# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class ServisPaylasimFormu:
    def __init__(self, katman_adi, servis_var, servis_ogc_uyumlu, servis_atlas_uyumlu, servis_wms_var, servis_wfs_var, servis_wms_version,
                servis_wfs_version, servis_aciklama, servis_yayin_platformu, sp_olmamasi_personel,
                sp_olmamasi_mevzuat, sp_olmamasi_donanim, sp_olmamasi_diger, sp_olmamasi_aciklama, adi, tucbs_katmani, inspire_katmani, xoid, k_adi):
        
        self.servis_turu = ''
        self.servis_version = ''
        self.katman_adi = katman_adi.rstrip()
        self.servis_var = servis_var
        self.servis_ogc_uyumlu = servis_ogc_uyumlu
        self.servis_atlas_uyumlu = servis_atlas_uyumlu
        self.servis_wms_var = servis_wms_var
        self.servis_wfs_var = servis_wfs_var
        self.servis_wms_version = servis_wms_version
        self.servis_wfs_version = servis_wfs_version
        self.servis_aciklama = servis_aciklama

        self.k_adi = k_adi

        if servis_yayin_platformu is not None:
            self.servis_yayin_platformu = cnn.getsinglekoddata('kod_ek_2_servis_platform', 'kod', 'objectid='+str(servis_yayin_platformu))
        else:
            self.servis_yayin_platformu = ''

        self.sp_olmamasi_personel = sp_olmamasi_personel
        self.sp_olmamasi_mevzuat = sp_olmamasi_mevzuat
        self.sp_olmamasi_donanim = sp_olmamasi_donanim
        if sp_olmamasi_diger is not None:
            self.sp_olmamasi_diger = sp_olmamasi_diger
        else:
            self.sp_olmamasi_diger = ''
        if sp_olmamasi_aciklama is not None:
            self.sp_olmamasi_aciklama = sp_olmamasi_aciklama
        else:
            self.sp_olmamasi_aciklama = ''
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
        elif self.servis_wms_var is False and self.servis_wfs_var:
            self.servis_turu = self.servis_turu + 'WFS'
        elif self.servis_wms_var and self.servis_wfs_var:
            self.servis_turu = self.servis_turu + 'WMS + WFS'
        else:
            self.servis_turu = ''
        if self.servis_wms_version is not None and self.servis_wfs_version is None:
            self.servis_version = self.servis_wms_version
        elif self.servis_wms_version is None and self.servis_wfs_version is not None:
            self.servis_version = self.servis_wfs_version
        elif self.servis_wms_version is not None and self.servis_wfs_version is not None:
            self.servis_version = 'WMS: '+self.servis_wms_version + '\n' + 'WFS: '+self.servis_wfs_version
        else:
            self.servis_version = ''
        self.xoid = xoid
        self.spoid = cnn.getsinglekoddata("x_ek_2_servis_paylasim", "count(objectid)", "x_ek_2_tucbs_veri_katmani="+str(xoid))
        
        if self.spoid > 0:
            self.servis_paylasilan_kurum = cnn.getlistofdata("x_ek_2_servis_paylasim", "*", "x_ek_2_tucbs_veri_katmani="+str(xoid))
        else:
            None


    def createExcelFile(self):
        try:

            excelPath = "created_excels"+"\\"+self.k_adi.decode('utf-8')+"\\"+u"CVSPMV"
            excelName = u"VSPAF.xlsx"
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
                except BaseException as ex:
                    print fullFolderPath
                    print ex

            # Dosya Olusturma
            workbook = xlsxwriter.Workbook(fullFolderPath+"\\"+excelName)
            worksheet = workbook.add_worksheet()

            # Satir / Sutun Ayarlari

            worksheet.set_column('A:A', 41.43)
            worksheet.set_column('B:B', 13.43)
            worksheet.set_column('C:C', 15)
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
            worksheet.set_row(21, 28.5)
            worksheet.set_row(22, 28.5)
            worksheet.set_row(23, 28.5)
            worksheet.set_row(24, 28.5)
            worksheet.set_row(26, 24)

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
            worksheet.write_rich_string('B1', main_header_format, u'Coğrafi Veri Servisleri ve Paylaşımı Analiz Formu', comment_format_c, u'\n(Her Tema İçin Ayrı Ayrı Doldurulacaktır.)', merge_format)
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
            worksheet.write_rich_string('A8', merge_format, u'Veri Katmanı',  comment_format_c, u'\n(Analiz Edilen Veri Katmanı Adı)', merge_format)
            worksheet.write_rich_string('B8', merge_format, u'Servis Durumu', comment_format_c, u'\n(Var/Yok)', merge_format)
            worksheet.write_rich_string('C8', merge_format, u'Servis Türü', comment_format_c, u'\n(WMS,WFS, vd.)', merge_format)
            worksheet.write('D8', u'Servis Versiyonu', merge_format)
            worksheet.merge_range('E8:F8', '', merge_format)
            worksheet.write_rich_string('E8', merge_format, u'OGC Uyumluluğu',  comment_format_c, u'\n(OGC Test linki eklenecek)', comment_format_c, '\n(Uyumlu/Uyumsuz)', merge_format)
            worksheet.merge_range('G8:H8', '', merge_format)
            worksheet.write_rich_string('G8', merge_format, u'ATLAS Uyumluluğu', comment_format_c, u'\n(Uyumlu/Uyumsuz)', merge_format)
            worksheet.merge_range('I8:K8', '', merge_format)
            worksheet.write_rich_string('I8', merge_format, u'Bağlantı Linki', comment_format_c, u'\n(Uyumluluk Testleri İçin Servis Erişim Linkleri)', merge_format)

            if self.katman_adi is not None:
                worksheet.write('A9', self.katman_adi.decode('utf-8'), data_format_c)
            else:
                worksheet.write('A9', u'', data_empty)
            
            if self.servis_var:
                worksheet.write('B9', u'Var', data_format_c)
                if self.servis_turu is not None:
                    worksheet.write_rich_string('C9', data_format_c, self.servis_turu.decode('utf-8'), data_format_c, '\n'+self.servis_yayin_platformu.decode('utf-8')+u' yayın platformu kullanılmaktadır', data_format_c)
                else:
                    worksheet.write('C9', u'', data_empty)
                if self.servis_version is not None:
                    worksheet.write('D9', self.servis_version.decode('utf-8'), data_format_c)
                else:
                    worksheet.write('D9', u'', data_empty)   
                if self.servis_ogc_uyumlu:
                    worksheet.merge_range('E9:F9', u'Uyumlu', data_format_c)
                else:
                    worksheet.merge_range('E9:F9', u'Bilinmiyor', data_format_c)
                if self.servis_atlas_uyumlu:
                    worksheet.merge_range('G9:H9', u'Uyumlu', data_format_c)
                else:
                    worksheet.merge_range('G9:H9', u'Bilinmiyor', data_format_c)
                worksheet.merge_range('I9:K9', u'-', data_format_c)

            else:
                worksheet.write('B9', u'Yok', data_format_c)
                worksheet.write('C9', '', border_format)
                worksheet.write('D9', '', border_format)
                worksheet.merge_range('E9:F9', '', border_format)
                worksheet.merge_range('G9:H9', '', border_format)
                worksheet.merge_range('I9:K9', '', border_format)
            
            

            worksheet.write('A10', u'', border_format)
            worksheet.write('A11', u'', border_format)
            worksheet.write('A12', u'', border_format)
            worksheet.write('A13', u'', border_format)
            worksheet.write('A14', u'', border_format)
            worksheet.write('A15', u'', border_format)
            worksheet.write('B10', u'', border_format)
            worksheet.write('B11', u'', border_format)
            worksheet.write('B12', u'', border_format)
            worksheet.write('B13', u'', border_format)
            worksheet.write('B14', u'', border_format)
            worksheet.write('B15', u'', border_format)
            worksheet.write('C10', u'', border_format)
            worksheet.write('C11', u'', border_format)
            worksheet.write('C12', u'', border_format)
            worksheet.write('C13', u'', border_format)
            worksheet.write('C14', u'', border_format)
            worksheet.write('C15', u'', border_format)
            worksheet.write('D10', u'', border_format)
            worksheet.write('D11', u'', border_format)
            worksheet.write('D12', u'', border_format)
            worksheet.write('D13', u'', border_format)
            worksheet.write('D14', u'', border_format)
            worksheet.write('D15', u'', border_format)
            worksheet.merge_range('E10:F10', u'', border_format)
            worksheet.merge_range('E11:F11', u'', border_format)
            worksheet.merge_range('E12:F12', u'', border_format)
            worksheet.merge_range('E13:F13', u'', border_format)
            worksheet.merge_range('E14:F14', u'', border_format)
            worksheet.merge_range('E15:F15', u'', border_format)
            worksheet.merge_range('G10:H10', u'', border_format)
            worksheet.merge_range('G11:H11', u'', border_format)
            worksheet.merge_range('G12:H12', u'', border_format)
            worksheet.merge_range('G13:H13', u'', border_format)
            worksheet.merge_range('G14:H14', u'', border_format)
            worksheet.merge_range('G15:H15', u'', border_format)
            worksheet.merge_range('I10:K10', u'', border_format)
            worksheet.merge_range('I11:K11', u'', border_format)
            worksheet.merge_range('I12:K12', u'', border_format)
            worksheet.merge_range('I13:K13', u'', border_format)
            worksheet.merge_range('I14:K14', u'', border_format)
            worksheet.merge_range('I15:K15', u'', border_format)
            worksheet.merge_range('A16:K16', u'', border_format)

            worksheet.write('A17', u'Servis Olmaması Durumu Açıklaması', text_format)
            if self.servis_aciklama is not None:
                h = 15*len(self.servis_aciklama)/60
                worksheet.merge_range('B17:K17', self.servis_aciklama.decode('utf-8'), data_format_r)
                if h < 28:
                    h = 28
                worksheet.set_row(16, h)
            else:
                worksheet.merge_range('B17:K17', u'(Servis Mevcut Değil İse Sebebi Burada Tariflenecektir)', comment_format_r)
            worksheet.write('A18', u'OGC Uyumlu Olmamasının Açıklaması', text_format)
            worksheet.merge_range('B18:K18', u'(OGC Uyumlu Veri Servisi Mevcut Değil İse Sebebi Burada Tariflenecektir)', comment_format_r)
            worksheet.write('A19', u'Atlas Uygulamasında Yayınlanamaması Durumu Açıklaması', text_format)
            worksheet.merge_range('B19:K19', u'(Atlas Uygulamasında Yayınamamasının Sebepleri Burada Tariflenecektir)', comment_format_r)
            worksheet.merge_range('A20:K20', '')

            # Servis Paylaşımı
            worksheet.merge_range('A21:K21', u'Servis Paylaşımı', header_format)
            worksheet.merge_range('A22:A25', '')
            if self.spoid > 0:
                worksheet.write_rich_string('A22', text_format, u'Servis Paylaşımı Var Mı? Evet(', data_format_c, u'X', merge_format, u') Hayır( )', merge_format)                
            else:
                worksheet.write_rich_string('A22', text_format, u'Servis Paylaşımı Var Mı? Evet( ) Hayır(', data_format_c, u'X', merge_format, u')', merge_format)
            worksheet.merge_range('B22:C25', u'Hayır ise Sebebi', merge_format)
            worksheet.merge_range('D22:K22', u'', merge_format)
            if self.sp_olmamasi_mevzuat:
                worksheet.write_rich_string('D22', text_format, u'Mevzuat Kaynaklı (', data_format_c, u'X', merge_format, u')', text_format)
            else:
                worksheet.write('D22', u'Mevzuat Kaynaklı ( )', text_format)
            worksheet.merge_range('D23:K23', u'', merge_format)
            if self.sp_olmamasi_personel:
                worksheet.write_rich_string('D23', text_format, u'Personel Kaynaklı (', data_format_c, u'X', merge_format, u')', text_format)
            else:
                worksheet.write('D23', u'Personel Kaynaklı ( )', text_format)
            worksheet.merge_range('D24:K24', u'', merge_format)
            if self.sp_olmamasi_donanim:
                worksheet.write_rich_string('D24', text_format, u'Donanım Kaynaklı (', data_format_c, u'X', merge_format, u')', text_format)
            else:
                worksheet.write('D24', u'Donanım Kaynaklı ( )', text_format)
            
            if (len(self.sp_olmamasi_diger) + len(self.sp_olmamasi_aciklama)) > 50:
                h = 15*len((self.sp_olmamasi_diger) + (self.sp_olmamasi_aciklama))/50
                if h < 28.5:
                    h = 28.5
                worksheet.set_row(14, h)
            worksheet.merge_range('D25:K25', self.sp_olmamasi_diger.decode('utf-8') + u' ' + self.sp_olmamasi_aciklama.decode('utf-8'), data_format_r)
            
            
            worksheet.write('A26', u'Evet ise Servis Paylaşılan Kurumlar', merge_format)
            worksheet.merge_range('B26:C26', u'Servis Durumu', merge_format)
            worksheet.merge_range('D26:K26', u'Protokol Durumu', merge_format)
            worksheet.write('A27', u'', merge_format)
            worksheet.write('B27', u'Ücretli ( )', merge_format)
            worksheet.write('C27', u'Ücretsiz ( )', merge_format)
            worksheet.merge_range('D27:G27', u'(Servis Paylaşımı İçin Yapılan Bir Protokol Var Mı? Var/Yok)', comment_format_c)
            worksheet.merge_range('H27:K27', u'(Var ise ek dosyalardan protokol örneği eklenecek.)', comment_format_c)
            
            last_starting_line = 28

            if self.spoid > 0:
                for i in self.servis_paylasilan_kurum:
                    paylasilan_kurum = i[2]
                    ucretli = i[3]
                    protokol_var = i[4]

                    if paylasilan_kurum is not None:
                        worksheet.write('A'+str(last_starting_line), paylasilan_kurum.decode('utf-8'), data_format_c)
                    else:
                        worksheet.write('A'+str(last_starting_line), u'', data_format_c)
                    if ucretli:
                        worksheet.write_rich_string('B'+str(last_starting_line), text_format, u'Ücretli (', data_format_c, u'X', merge_format, u')', merge_format)
                        worksheet.write('C'+str(last_starting_line), u'Ücretsiz ( )', merge_format)
                    else:
                        worksheet.write('B'+str(last_starting_line), u'Ücretli ( )', merge_format)
                        worksheet.write_rich_string('C'+str(last_starting_line), merge_format, u'Ücretsiz (', data_format_c, u'X', merge_format, ')', merge_format)
                    if protokol_var:
                        worksheet.merge_range('D'+str(last_starting_line)+':G'+str(last_starting_line), 'Var', data_format_c)
                    else:
                        worksheet.merge_range('D'+str(last_starting_line)+':G'+str(last_starting_line), 'Yok', data_format_c)
                    worksheet.merge_range('H'+str(last_starting_line)+':K'+str(last_starting_line), '', merge_format)
                    last_starting_line += 1

            # Calisma Grubu
            worksheet.merge_range('A'+str(last_starting_line)+':K'+str(last_starting_line), '')
            last_starting_line += 1
            worksheet.merge_range('A'+str(last_starting_line)+':K'+str(last_starting_line), u'Çalışma Grubu', header_format)            
            last_starting_line += 1
            worksheet.write('A'+str(last_starting_line), u'Grup Adı', text_format)            
            worksheet.merge_range('B'+str(last_starting_line)+':K'+str(last_starting_line), u'(Analiz Grubu Adı)', comment_format_r)
            last_starting_line += 1
            worksheet.write('A'+str(last_starting_line), u'Adı Soyadı', text_format)
            worksheet.merge_range('B'+str(last_starting_line)+':K'+str(last_starting_line), u'(Analizi Gerçekleştiren Proje Uzmanı)', comment_format_r)
            last_starting_line += 1
            worksheet.write('A'+str(last_starting_line), u'Görevi', text_format)
            worksheet.merge_range('B'+str(last_starting_line)+':K'+str(last_starting_line), u'(Proje Uzmanı Görevi)', comment_format_r)
            last_starting_line += 1
            worksheet.write('A'+str(last_starting_line), u'Tarih', text_format)
            worksheet.merge_range('B'+str(last_starting_line)+':K'+str(last_starting_line), u'(Analiz Tarihi)', comment_format_r)
            
            workbook.close()
        except Exception as e:
            import sys
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
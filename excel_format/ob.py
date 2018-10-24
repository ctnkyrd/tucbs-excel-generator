# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class OrganizasyonBirimiFormu:
    def __init__(self,bakanlik, adi, k_adi, fileCount, cbs_birimi_var, cbs_yeni_birim_gerekli, veri_uretim_var, tasra_teskilati_var, sadece_sistem_idamesi_var,
                    ihtiyaclar_sizin_onayinizdan_geciyor, personel_yeterli, personel_yetersizlik_kriterleri, personel_yetersizlik_oneri, 
                    cbs_farkindaligi_var, sorunlar, cbs_birim_adi, yapilanma_olcegi, kurum_semasindaki_yeri, tasra_teskilat_yapilanmasi, cbs_yeni_birim_gorusler
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
        
        # if personel_yetersizlik_kriterleri is not None:
        #     self.personel_yetersizlik_kriterleri = cnn.getsinglekoddata('kod_ek_cbs_org_personel_yetersizlik', 'kod', 'objectid='+str(personel_yetersizlik_kriterleri))
        # else:
        #     self.personel_yetersizlik_kriterleri = None

        self.personel_yetersizlik_kriterleri = personel_yetersizlik_kriterleri

        self.personel_yetersizlik_oneri = personel_yetersizlik_oneri
        self.cbs_farkindaligi_var = cbs_farkindaligi_var
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

            just_border = wb.add_format()
            just_border.set_border()
            just_border.set_align('vcenter')
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
            
            header_format_left = wb.add_format({
                'font_size': 16,
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'})

            sub_header_format = wb.add_format({
               'font_size': 11,
                'bold': 1,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter' 
            })
            sub_header_format.set_text_wrap()

            light_header_format = wb.add_format({
               'font_size': 11,
                'border': 1,
                'align': 'left',
                'valign': 'vcenter' 
            })

            header_format = wb.add_format({
                'font_size': 12,
                'bold': 1,
                'border': 1,
                'align': 'center',
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
            ws.set_column('G:G', 19.86)
            ws.set_column('H:H', 32.71)

            
            ws.merge_range('A1:A4', u'', main_header_format)
            ws.merge_range('A5:H5', u'', merge_format)
            ws.merge_range('A17:H17', u'', merge_format)
            ws.merge_range('H1:H4', u'', main_header_format)

            ws.merge_range('B1:F4', u'CBS Organizasyon Birimleri ve İnsan Kaynakları Analiz Formu', main_header_format)
            ws.write('G1', u'Revizyon Numarası', merge_format)
            ws.write('G3', u'Revizyon Tarihi', merge_format)
            ws.write('G4', u'', merge_format)
            ws.merge_range('C10:H11', u'', merge_format)



            # Baslik
            ws.insert_image('A1', r"logo\csb.jpg", {'x_offset': 70,'y_offset': 5,'x_scale': 1.25})
            ws.insert_image('H1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})

            ws.merge_range('A6:H6', u'Genel Bilgiler', header_format_left)
            ws.write('A7', u'Bakanlık', sub_header_format)
            ws.write('A8', u'Genel Müdürlük / Belediye', sub_header_format)
            ws.write('A9', u'CBS Birimi Var Mı?', sub_header_format)
            ws.merge_range('A10:A11', u'CBS ile İlgili Yeni Bir Birim Kurma Gereksinimi Var Mı? ', sub_header_format)
            ws.write('A12', u'CBS Birim Adı', sub_header_format)
            ws.write('A13', u'Hangi lçekte yapılandırılmış?', sub_header_format)
            ws.write('A14', u'Kurum şemasındaki yeri', sub_header_format)
            ws.write('A15', u'Taşra teşkilatı var mı?', sub_header_format)
            ws.write('A16', u'Taşra teşkilatı yapılanması', sub_header_format)
            ws.merge_range('A18:H18', u'Görev Bilgileri', header_format_left)
            ws.write('A19', u'Veriyi Üreten Birim', sub_header_format)
            ws.write('A20', u'Veriyi Sunan Birim', sub_header_format)
            ws.set_row(20, 30)
            ws.write('A21', u'CBS ile ilgili her ihtiyaç sizin onay ve kontrolünüzden mi geçiyor?', sub_header_format)
            ws.set_row(21, 8.25)
            ws.write('A23', u'CBS birim personeli yeterli mi?', sub_header_format)
            ws.write('A24', u'Personelin yetersizlik kriterleri', sub_header_format)
            ws.set_row(24, 30)
            ws.write('A25', u'Yönetim Düzeyinde CBS Farkındalığı Var mı?', sub_header_format)
            ws.set_row(25, 30)
            ws.write('A26', u'Sorunlar', sub_header_format)

            ws.merge_range('B7:H7', self.bakanlik, data_format_r)
            ws.merge_range('B8:H8', self.adi, data_format_r)
            if self.cbs_birimi_var:
                ws.merge_range('B9:H9', u'Evet', data_format_r)
            else:
                ws.merge_range('B9:H9', u'Hayır', data_format_r)
            if self.cbs_yeni_birim_gerekli:
                ws.write_rich_string('B10',light_header_format, u'Evet ( ', data_format_c,u'X', light_header_format, u' )')
                ws.write('B11', u'Hayır (    )',light_header_format)
            else:
                ws.write_rich_string('B11',light_header_format, u'Hayır ( ', data_format_c,u'X', light_header_format, u' )')
                ws.write('B10', u'Evet (    )',light_header_format)
            

            if self.cbs_yeni_birim_gorusler is not None:
                ws.write('C10', self.cbs_yeni_birim_gorusler.decode('utf-8'), data_format_r)

            
            if self.cbs_birim_adi is not None:
                ws.merge_range('B12:H12', self.cbs_birim_adi.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B12:H12', u'', data_format_r)
            
            if self.yapilanma_olcegi is not None:
                ws.merge_range('B13:H13', self.yapilanma_olcegi.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B13:H13', u'', data_format_r)
            
            if self.kurum_semasindaki_yeri is not None:
                ws.merge_range('B14:H14', self.kurum_semasindaki_yeri.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B14:H14', u'', data_format_r)

            if self.tasra_teskilati_var:
                ws.merge_range('B15:H15', u'Evet', data_format_r)
            else:
                ws.merge_range('B15:H15', u'Hayır', data_format_r)
            
            if self.tasra_teskilat_yapilanmasi is not None:
                ws.merge_range('B16:H16', self.tasra_teskilat_yapilanmasi.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B16:H16', u'', data_format_r)

            if self.veri_uretim_var is not None:
                ws.merge_range('B19:H19', self.veri_uretim_var.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B19:H19', u'', data_format_r)

            if self.sadece_sistem_idamesi_var is not None:
                ws.merge_range('B20:H20', self.sadece_sistem_idamesi_var.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B20:H20', u'', data_format_r)

            if self.ihtiyaclar_sizin_onayinizdan_geciyor:
                ws.merge_range('B21:H21', u'Evet', data_format_r)
            else:
                ws.merge_range('B21:H21', u'Hayır', data_format_r)

            if self.personel_yeterli:
                ws.write_rich_string('B23', light_header_format, u'Evet ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('C23', u'Hayır (   )' ,light_header_format)
            else:
                ws.write_rich_string('C23', light_header_format, u'Hayır ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('B23', u'Evet (   )' ,light_header_format)

            ws.merge_range('D23:H23', u'Öneriler', header_format)


            if self.personel_yetersizlik_kriterleri == 1:
                ws.write_rich_string('B24', light_header_format, u'Sayı ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('C24', u'Nitelik (   )' ,light_header_format)
            elif self.personel_yetersizlik_kriterleri == 2:
                ws.write_rich_string('C24', light_header_format, u'Nitelik ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('B24', u'Sayı (   )' ,light_header_format)
            else:
                ws.write('B24', u'Sayı (   )' ,light_header_format)
                ws.write('C24', u'Nitelik (   )' ,light_header_format)
            
            if self.cbs_farkindaligi_var:
                ws.write_rich_string('B25', light_header_format, u'Evet ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('C25', u'Hayır (   )' ,light_header_format)
            else:
                ws.write_rich_string('C25', light_header_format, u'Hayır ( ', data_format_c,u'X', light_header_format, u' )',just_border)
                ws.write('B25', u'Evet (   )' ,light_header_format)
            
            ws.merge_range('D24:H25', u'', light_header_format)

            if self.personel_yetersizlik_oneri is not None:
                ws.write('D24', self.personel_yetersizlik_oneri.decode('utf-8'), data_format_r)
            
            if self.sorunlar is not None:
                ws.merge_range('B26:H26', self.sorunlar.decode('utf-8'), data_format_r)
            else:
                ws.merge_range('B26:H26', u'', data_format_c)
                

                

            



            



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
        except Exception as e:
            import sys
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
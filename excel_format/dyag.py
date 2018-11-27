# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class DonanimYazilimFormu:
    def __init__(self, bakanlik, adi, birim, fileCount, sunucu_yeterli, sunucu_yetersiz_aciklama, kamunet_agina_bagli, kamunet_agina_bagli_degil_aciklama, 
                    ipsecvpn_uygun, ipsecvpn_uygunsuz_aciklama, ipsecvpn_bagli, k_adi):

        self.bakanlik = bakanlik
        self.adi = adi.rstrip()
        self.k_adi = k_adi
        self.birim = birim
        self.fileCount = fileCount
        self.sunucu_yeterli = sunucu_yeterli
        self.sunucu_yetersiz_aciklama = sunucu_yetersiz_aciklama
        self.kamunet_agina_bagli = kamunet_agina_bagli
        self.kamunet_agina_bagli_degil_aciklama = kamunet_agina_bagli_degil_aciklama
        self.ipsecvpn_uygun = ipsecvpn_uygun
        self.ipsecvpn_uygunsuz_aciklama = ipsecvpn_uygunsuz_aciklama
        self.ipsecvpn_bagli = ipsecvpn_bagli


    def createExcelFile(self):
        try:
            excelPath = "created_excels"+"\\"+self.k_adi.decode('utf-8')
            if self.fileCount == 0:
                excelName = u"DYAGAF.xlsx"
            else:
                excelName = u"DYAGAF_"+str(self.fileCount)+".xlsx"

            fullFolderPath = excelPath

            if os.path.isdir(unicode(fullFolderPath)) is False:
                try:
                    os.makedirs(unicode(fullFolderPath))
                except BaseException as ex:
                    print fullFolderPath
                    print ex
            
            wb = xlsxwriter.Workbook(fullFolderPath+"\\"+excelName)
            # Dosya Olusturma
            worksheet = wb.add_worksheet()

            # Satir / Sutun Ayarlari
            worksheet.set_column('A:A', 34.57)
            worksheet.set_column('B:B', 14.14)
            worksheet.set_column('C:C', 19.14)
            worksheet.set_column('D:D', 18.86)
            worksheet.set_column('E:E', 21)
            worksheet.set_column('F:F', 22.14)
            worksheet.set_column('G:G', 25.57)

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

            f_data_emty = wb.add_format()
            f_data_emty.set_bg_color('#C5C5C5')
            f_data_emty.set_border()

            default_format = wb.add_format()
            default_format.set_border()

            light_format = wb.add_format()
            light_format.set_border()

            f_comment_left = wb.add_format()
            f_comment_left.set_border()
            f_comment_left.set_color('gray')
            f_comment_left.set_italic()
            f_comment_left.set_font_size(9)
            f_comment_left.set_valign('vcenter')

            # Baslik
            worksheet.insert_image('A1', r"logo\csb.jpg", {'x_offset': 43,'y_offset': 7,'x_scale': 1.5})
            worksheet.merge_range('A1:A4', u'', merge_format)
            worksheet.merge_range('B1:E4', u'Donanım, Yazılım, Ağ ve Güvenlik Analiz Formu', main_header_format)
            worksheet.write('F1', u'Revizyon Numarası', merge_format)
            worksheet.write('F3', u'Revizyon Tarihi', merge_format)
            worksheet.write('F4', u'', merge_format)
            worksheet.insert_image('G1', r"logo\tucbs2.jpg", {'x_offset': 9,'y_offset': 9,'x_scale': 0.43,'y_scale': 0.43})
            worksheet.merge_range('G1:G4', u'', merge_format)
            worksheet.merge_range('A5:G5', u'')

            # Genel Bilgiler
            worksheet.merge_range('A6:G6', u'Genel Bilgiler', header_format)
            worksheet.write('A7', u'Bakanlık', text_format)

            if self.bakanlik is not None:
                worksheet.merge_range('B7:G7', self.bakanlik, data_format_r)
            else:
                worksheet.merge_range('B7:G7', u'', f_data_emty)
            
            if self.adi is not None:
                worksheet.merge_range('B8:G8', self.adi, data_format_r)
            else:
                worksheet.merge_range('B8:G8', u'', f_data_emty)

            if self.birim is not None:
                worksheet.merge_range('B9:G9', self.birim, data_format_r)
            else:
                worksheet.merge_range('B9:G9', u'', f_data_emty)
                

            worksheet.write('A8', u'Genel Müdürlük / Belediye', text_format)
            worksheet.write('A9', u'Birimi Adı', text_format)
            worksheet.merge_range('A10:G10', u'')


            # Donanim
            worksheet.merge_range('A11:G11', u'Donanım', header_format)
            worksheet.merge_range('A12:A13', u'Coğrafi Veri Depolama ve Sunumu Amaçlı Kullanılan Donanım Yeterli Mi?', text_format)

            if self.sunucu_yeterli:
                worksheet.write_rich_string('B12', light_format, u'Evet ( ', data_format_r, u'X', light_format, u' )')
                worksheet.write('B13', u'Hayır ( )', light_format)
            else:
                worksheet.write('B12', u'Evet ( )', light_format)
                worksheet.write_rich_string('B13', light_format, u'Hayır ( ', data_format_r, u'X', light_format, u' )',default_format)

            if self.sunucu_yetersiz_aciklama is not None:
                worksheet.merge_range('C12:G13', self.sunucu_yetersiz_aciklama.decode('utf-8').rstrip(), data_format_r)
            else:
                worksheet.merge_range('C12:G13', u'', data_format_r)

            worksheet.merge_range('A14:G14', u'')

            # Ag ve Guvenlik
            worksheet.merge_range('A15:G15', u'Ağ ve Güvenlik', header_format)
            worksheet.merge_range('A16:A17', u'Kamu.Net Ağına Bağlı', text_format)

            if self.kamunet_agina_bagli:
                worksheet.write_rich_string('B16', light_format, u'Evet ( ', data_format_r, u'X', light_format, u' )')
                worksheet.write('B17', u'Hayır ( )', light_format)
            else:
                worksheet.write('B16', u'Evet ( )', light_format)
                worksheet.write_rich_string('B17', light_format, u'Hayır ( ', data_format_r, u'X', light_format, u' )',default_format) 

            if self.kamunet_agina_bagli_degil_aciklama is not None:
                worksheet.merge_range('C16:G17', self.kamunet_agina_bagli_degil_aciklama.decode('utf-8'), data_format_r)
            else:
                worksheet.merge_range('C16:G17', u'', data_format_r)


            worksheet.merge_range('A18:A19', u'IPSECVPN Olarak Bağlantı Yapmaya Uygun Mu?', text_format)


            if self.ipsecvpn_uygun:
                worksheet.write_rich_string('B18', light_format, u'Evet ( ', data_format_r, u'X', light_format, u' )',default_format)
                worksheet.write('B19', u'Hayır ( )', light_format)
            else:
                worksheet.write('B18', u'Evet ( )', light_format)
                worksheet.write_rich_string('B19', light_format, u'Hayır ( ', data_format_r, u'X', light_format, u' )', default_format) 

            if self.ipsecvpn_uygunsuz_aciklama is not None:
                worksheet.merge_range('C18:G19', self.ipsecvpn_uygunsuz_aciklama.decode('utf-8'), data_format_r)
            else:
                worksheet.merge_range('C18:G19', u'', data_format_r)
            worksheet.merge_range('A20:G20', u'(Bu Form Çevre ve Şehircilik Bakanlığı Ağ ve Güvenlik Kuralları İdareden Temin Edildikten Sonra Revize Edilmelidir)', f_comment_left)

            wb.close()
        except BaseException as ex:
            print ex
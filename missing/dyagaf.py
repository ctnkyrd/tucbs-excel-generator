# -*- coding: utf-8 -*-
import xlsxwriter

# Dosya Olusturma
workbook = xlsxwriter.Workbook('dyagaf.xlsx')
worksheet = workbook.add_worksheet()

# Satir / Sutun Ayarlari
worksheet.set_column('A:A', 34.57)
worksheet.set_column('B:B', 14.14)
worksheet.set_column('C:C', 19.14)
worksheet.set_column('D:D', 18.86)
worksheet.set_column('E:E', 21)
worksheet.set_column('F:F', 22.14)
worksheet.set_column('G:G', 25.57)

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
worksheet.merge_range('B7:G7', u'İçişleri Bakanlığı', data_format_r)
worksheet.write('A8', u'Genel Müdürlük / Belediye', text_format)
worksheet.merge_range('B8:G8', u'Afet ve Acil Durum Yönetimi Başkanlığı', data_format_r)
worksheet.write('A9', u'Birimi Adı', text_format)
worksheet.merge_range('B9:G9', u'Coğrafi Bilgi Teknolojileri Çalışma Grubu', data_format_r)
worksheet.merge_range('A10:G10', u'')

# Donanim
worksheet.merge_range('A11:G11', u'Donanım', header_format)
worksheet.merge_range('A12:A13', u'Coğrafi Veri Depolama ve Sunumu Amaçlı Kullanılan Donanım Yeterli Mi?', text_format)
worksheet.write('B12', u'Evet ()', text_format)
worksheet.write('B13', u'Hayır ()', text_format)
worksheet.merge_range('C12:G13', u'', data_format_r)
worksheet.merge_range('A14:G14', u'')

# Ag ve Guvenlik
worksheet.merge_range('A15:G15', u'Ağ ve Güvenlik', header_format)
worksheet.merge_range('A16:A17', u'Kamu.Net Ağına Bağlı', text_format)
worksheet.write('B16', u'Evet ()', text_format)
worksheet.write('B17', u'Hayır ()', text_format)
worksheet.merge_range('C16:G17', u'', data_format_r)
worksheet.merge_range('A18:A19', u'IPSECVPN Olarak Bağlantı Yapmaya Uygun Mu?', text_format)
worksheet.write('B18', u'Evet ()', text_format)
worksheet.write('B19', u'Hayır ()', text_format)
worksheet.merge_range('C18:G19', u'', data_format_r)

workbook.close()
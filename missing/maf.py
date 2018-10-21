# -*- coding: utf-8 -*-
import xlsxwriter

# Dosya Olusturma
workbook = xlsxwriter.Workbook('maf.xlsx')
worksheet = workbook.add_worksheet()

# Satir / Sutun Ayarlari
worksheet.set_column('A:A', 30)
worksheet.set_column('B:B', 22)
worksheet.set_column('C:C', 20)
worksheet.set_column('D:D', 8.5)
worksheet.set_column('E:E', 22)
worksheet.set_column('F:F', 26)

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
worksheet.insert_image('A1', r"logo\csb.jpg", {'x_offset': 35,'y_offset': 7,'x_scale': 1.4})
worksheet.merge_range('A1:A4', '', merge_format)
worksheet.merge_range('B1:D4', 'Mevzuat Analiz Formu', main_header_format)
worksheet.write('E1', u'Revizyon Numarası', merge_format)
worksheet.write('E3', 'Revizyon Tarihi', merge_format)
worksheet.write('E4', '', merge_format)
worksheet.insert_image('F1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})
worksheet.merge_range('F1:F4', '', merge_format)
worksheet.merge_range('A5:F5', '')

# Genel Bilgiler
worksheet.merge_range('A6:F6', 'Genel Bilgiler', header_format)
worksheet.write('A7', u'Bakanlık', text_format)
worksheet.merge_range('B7:F7', 'Icisleri Bakanligi', data_format_r)
worksheet.write('A8', u'Genel Müdürlük / Belediye', text_format)
worksheet.merge_range('B8:F8', 'Afet ve Acil Durum Yonetimi Baskanligi', data_format_r)
worksheet.write('A9', u'Birim Adı', text_format)
worksheet.merge_range('B9:F9', 'Cografi Bilgi Teknolojileri Calisma Grubu', data_format_r)
worksheet.merge_range('A10:F10', '')
worksheet.write('A11', u'Metaveri ve Coğrafi Veri Servis Paylaşımı ile ilgili mevzuat hakkında bir kısıtlama var mı?', text_format)
worksheet.merge_range('B11:F11', 'Veri paylasimina engel mevzuatsal bir kisit bulunmamaktadir ancak Geoportal ile paylasilabilecek verilerin yonetim tarafindan belirlenmesi gerekmektedir.', data_format_r)
worksheet.merge_range('A12:F12', '')

# Ilgili Mevzuat
worksheet.merge_range('A13:F13', u'İlgili Mevzuat', header_format)
worksheet.write('A14', u'Adı / Numarası', merge_format)
worksheet.write('B14', u'İlgili Maddeler', merge_format)
worksheet.write('C14', u'İlişkili Olduğu Süreç', merge_format)
worksheet.merge_range('D14:E14', u'Veri Paylaşımına Etkisi', merge_format)
worksheet.write('F14', u'Etkilediği Tema / Katman', merge_format)
worksheet.write('A15', '', data_format_c)
worksheet.write('A16', '', data_format_c)
worksheet.write('A17', '', data_format_c)
worksheet.write('A18', '', data_format_c)
worksheet.write('A19', '', data_format_c)
worksheet.write('A20', '', data_format_c)
worksheet.write('B15', '', data_format_c)
worksheet.write('B16', '', data_format_c)
worksheet.write('B17', '', data_format_c)
worksheet.write('B18', '', data_format_c)
worksheet.write('B19', '', data_format_c)
worksheet.write('B20', '', data_format_c)
worksheet.write('C15', '', data_format_c)
worksheet.write('C16', '', data_format_c)
worksheet.write('C17', '', data_format_c)
worksheet.write('C18', '', data_format_c)
worksheet.write('C19', '', data_format_c)
worksheet.write('C20', '', data_format_c)
worksheet.write('D15', '', data_format_c)
worksheet.write('D16', '', data_format_c)
worksheet.write('D17', '', data_format_c)
worksheet.write('D18', '', data_format_c)
worksheet.write('D19', '', data_format_c)
worksheet.write('D20', '', data_format_c)
worksheet.write('E15', '', data_format_c)
worksheet.write('E16', '', data_format_c)
worksheet.write('E17', '', data_format_c)
worksheet.write('E18', '', data_format_c)
worksheet.write('E19', '', data_format_c)
worksheet.write('E20', '', data_format_c)
worksheet.write('F15', '', data_format_c)
worksheet.write('F16', '', data_format_c)
worksheet.write('F17', '', data_format_c)
worksheet.write('F18', '', data_format_c)
worksheet.write('F19', '', data_format_c)
worksheet.write('F20', '', data_format_c)
worksheet.write('A21', u'Coğrafi Veri Paylaşılamama Sebebi', text_format)
worksheet.merge_range('B21:F21', u'', data_format_r)



workbook.close()
# -*- coding: utf-8 -*-
import xlsxwriter

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
worksheet.merge_range('E8:P8', u'Afete Maruz Bölgeler', data_format_r)
worksheet.merge_range('A9:D9', u'Metaveri Var Mı?', text_format)
worksheet.merge_range('E9:P9', u'Hayır', data_format_r)
worksheet.merge_range('A10:D10', u'Metaveri Hangi Standarta Uygun Üretiliyor?', text_format)
worksheet.merge_range('E10:P10', u'', data_format_r)
worksheet.merge_range('A11:D11', u'Metaveri Yayınlanıyor Mu?', text_format)
worksheet.merge_range('E11:P11', u'Hayır', data_format_r)
worksheet.merge_range('A12:D12', u'CBS Genel Müdürlüğü ile Paylaşımı Var Mı? ', text_format)
worksheet.merge_range('E12:P12', u'Hayır', data_format_r)
worksheet.merge_range('A13:D13', u'Açıklama', text_format)
worksheet.merge_range('E13:P13', u'', data_format_r)

workbook.close()
# -*- coding: utf-8 -*-
import xlsxwriter

# Dosya Olusturma
workbook = xlsxwriter.Workbook('obaf.xlsx')
worksheet = workbook.add_worksheet()

# Satir / Sutun Ayarlari
worksheet.set_column('A:A', 38.29)
worksheet.set_column('B:B', 13.71)
worksheet.set_column('C:C', 14)
worksheet.set_column('D:D', 15.86)
worksheet.set_column('E:E', 17.29)
worksheet.set_column('F:F', 17.57)
worksheet.set_column('G:G', 19.86)
worksheet.set_column('H:H', 32.71)
worksheet.set_row(25, 42)

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
worksheet.insert_image('A1', r"D:\Github\tucbs-excel-generator\logo\csb.jpg", {'x_offset': 50,'y_offset': 7,'x_scale': 1.5})
worksheet.merge_range('A1:A4', '', merge_format)
worksheet.merge_range('B1:F4', u'CBS Organizasyon Birimleri ve İnsan Kaynakları Analiz Formu', main_header_format)
worksheet.write('G1', u'Revizyon Numarası', merge_format)
worksheet.write('G3', 'Revizyon Tarihi', merge_format)
worksheet.write('G4', '', merge_format)
worksheet.insert_image('H1', r"D:\Github\tucbs-excel-generator\logo\tucbs2.jpg", {'x_offset': 30,'y_offset': 7,'x_scale': 0.45,'y_scale': 0.45})
worksheet.merge_range('H1:H4', '', merge_format)
worksheet.merge_range('A5:H5', '')

# Genel Bilgiler
worksheet.merge_range('A6:H6', 'Genel Bilgiler', header_format)
worksheet.write('A7', u'Bakanlık', text_format)
worksheet.merge_range('B7:H7', u'İçişleri Bakanlığı', data_format_r)
worksheet.write('A8', u'Genel Müdürlük / Belediye', text_format)
worksheet.merge_range('B8:H8', u'Afet ve Acil Durum Yönetimi Başkanlığı', data_format_r)
worksheet.write('A9', u'CBS Birimi Var Mı?', text_format)
worksheet.merge_range('B9:H9', u'Var', data_format_r)
worksheet.merge_range('A10:A11', u'CBS ile İlgili Yeni Bir Birim Kurma Gereksinimi Var Mı? ', text_format)
worksheet.write('B10', u'Evet ()', text_format)
worksheet.write('B11', u'Hayır ()', text_format)
worksheet.merge_range('C10:H11', u'', data_format_r)
worksheet.write('A12', u'CBS Birim Adı', text_format)
worksheet.merge_range('B12:H12', u'Coğrafi Bilgi Teknolojileri Çalışma Grubu', data_format_r)
worksheet.write('A13', u'Hangi Ölçekte Yapılandırılmış', text_format)
worksheet.merge_range('B13:H13', u'Çalışma Grubu', data_format_r)
worksheet.write('A14', u'Kurum Şemasındaki Yeri', text_format)
worksheet.merge_range('B14:H14', u'Bilgi Sisteöleri ve Haberleşme Dairesi Başkanlığı Altında', data_format_r)
worksheet.write('A15', u'Taşra Teşkilatı Var Mı?', text_format)
worksheet.merge_range('B15:H15', u'', data_format_r)
worksheet.write('A16', u'Taşra Teşkilatı Yapılanması', text_format)
worksheet.merge_range('B16:H16', u'', data_format_r)
worksheet.merge_range('A17:H17', '')

# Gorev Bilgileri
worksheet.merge_range('A18:H18', u'Görev Bilgileri', header_format)
worksheet.write('A19', u'Veriyi Üreten Birim', text_format)
worksheet.merge_range('B19:H19', u'İlgili Daire Başkanlıkları', data_format_r)
worksheet.write('A20', u'Veriyi Sunan Birim', text_format)
worksheet.merge_range('B20:H20', u'İçişleri Bakanlığı', data_format_r)
worksheet.write('A21', u'CBS ile İlgili Her İhtiyaç Sizin Onay ve Kontrolünüzden Mi Geçiyor?', text_format)
worksheet.merge_range('B21:H21', u'Evet', data_format_r)
worksheet.merge_range('A22:H22', '')
worksheet.write('A23', u'CBS Birimi Personeli Yeterli Mi?', text_format)
worksheet.write('B23', u'Evet ()', text_format)
worksheet.write('C23', u'Hayır ()', text_format)
worksheet.merge_range('D23:H23', u'Öneriler', text_format)
worksheet.write('A24', u'Personelin Yetersizlik Kriterleri', text_format)
worksheet.write('B24', u'Sayı ()', text_format)
worksheet.write('C24', u'Nitelik ()', text_format)
worksheet.write('A25', u'Yönetim Düzeyinde CBS Farkındalığı Var Mı?', text_format)
worksheet.write('B25', u'Evet ()', text_format)
worksheet.write('C25', u'Hayır ()', text_format)
worksheet.merge_range('D24:H25', u'', data_format_r)
worksheet.write('A26', u'Sorunlar', text_format)
worksheet.merge_range('B26:H26', u'Personel sayı ve nitelik açısından yetersizdir.', data_format_r)



workbook.close()
# -*- coding: utf-8 -*-
import xlsxwriter, os
from pgget import Connection
cnn = Connection()

class MevzuatAnalizFormu:
    def __init__(self,bakanlik, adi, k_adi, birim, fileCount, oid, mevzuat_kisitlama, veri_paylasmama_sebep):
        # from kurum clas
        self.bakanlik = bakanlik
        self.adi = adi
        self.k_adi = k_adi
        self.birim = birim
        self.fileCount = fileCount

        # from table
        self.oid = oid
        self.mevzuat_kisitlama = mevzuat_kisitlama
        self.veri_paylasmama_sebep = veri_paylasmama_sebep

        self.mevzuat_sayisi = cnn.getsinglekoddata('x_ek_mevzuat_analiz_ilgili_mevzuat', 'count(*)', 'ek_mevzuat='+str(oid))
        if self.mevzuat_sayisi > 0:
            self.ilgili_mevzuat = cnn.getlistofdata('x_ek_mevzuat_analiz_ilgili_mevzuat','*', 'ek_mevzuat='+str(oid))
        else:
            self.ilgili_mevzuat = None
        

        
    def createExcelFile(self):
        try:

            excelPath = "created_excels"+"\\"+self.k_adi.decode('utf-8')
            if self.fileCount == 0:
                excelName = u"MAF.xlsx"
            else:
                excelName = u"MAF_"+str(self.fileCount)+".xlsx"

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

            ws.set_column('A:A', 30)
            ws.set_column('B:B', 22.57)
            ws.set_column('C:C', 20.29)
            ws.set_column('D:D', 8.71)
            ws.set_column('E:E', 21.86)
            ws.set_column('F:F', 25.86)

            
            ws.merge_range('A1:A4', u'', main_header_format)
            ws.merge_range('A5:F5', u'')
            ws.merge_range('A10:F10', u'')
            ws.merge_range('A12:F12', u'')
            ws.merge_range('F1:F4', u'', main_header_format)


            ws.merge_range('B1:D4', u'Mevzuat Analiz Formu', main_header_format)
            ws.write('E1', u'Revizyon Numarası', merge_format)
            ws.write('E3', u'Revizyon Tarihi', merge_format)
            ws.write('E4', u'', merge_format)


            # Baslik
            ws.insert_image('A1', r"logo\csb.jpg", {'x_offset': 70,'y_offset': 5,'x_scale': 1.25})
            ws.insert_image('F1', r"logo\tucbs2.jpg", {'x_offset': 6,'y_offset': 7,'x_scale': 0.44,'y_scale': 0.44})

            ws.merge_range('A6:F6', u'Genel Bilgiler', header_format_left)
            ws.write('A7', u'Bakanlık', sub_header_format)
            ws.write('A8', u'Genel Müdürlük / Belediye', sub_header_format)
            ws.write('A9', u'Birim Adı', sub_header_format)

            ws.merge_range('B7:F7', self.bakanlik, data_format_r)
            ws.merge_range('B8:F8', self.adi, data_format_r)
            ws.merge_range('B9:F9', self.birim, data_format_r)


            ws.set_row(10, 45.75)
            ws.set_row(5, 21)
            ws.set_row(12, 21)


            ws.write('A11', u'Metaveri ve Coğrafi Veri Servis Paylaşımı ile ilgili mevzuat hakkında bir kısıtlama varmı?', sub_header_format)

            ws.merge_range('A13:F13', u'İlgili Mevzuat', header_format_left)
            ws.write('A14', u'Adı/Numarası', header_format)
            ws.write('B14', u'İlgili Maddeler', header_format)
            ws.write('C14', u'İlişkili Olduğu Süreç', header_format)
            ws.merge_range('D14:E14', u'Veri Paylaşımına Etkileri', header_format)
            ws.write('F14', u'Etkilediği Tema/Katman', header_format)

            if self.mevzuat_kisitlama:
                ws.merge_range('B11:F11', u'Evet. Kurum tarafından veri paylaşımına engel mevzuat kısıtlamasının olduğu ifade edilmiştir.', data_format_r)
            else:
                ws.merge_range('B11:F11', u'Hayır. Kurum tarafından veri paylaşımına engel herhangi bir mevzuat kısıtlamasının bulunmadığı ifade edilmiştir.', data_format_r)

            last_starting_line = 15

            if self.mevzuat_sayisi > 0:
                for i in self.ilgili_mevzuat:
                    adi_numarasi = i[2]
                    ilgili_maddeler = i[11]
                    iliskili_surec = i[12]
                    veri_paylasimina_etkileri = i[13]
                    etkiledigi_tema_katman = i[14]
                    
                    if adi_numarasi is not None:
                        ws.write('A'+str(last_starting_line), adi_numarasi.decode('utf-8'), data_format_c)
                    else:
                        ws.write('A'+str(last_starting_line), u'', data_format_c)
                    if ilgili_maddeler is not None:
                        ws.write('B'+str(last_starting_line), ilgili_maddeler.decode('utf-8'), data_format_c)
                    else:
                        ws.write('B'+str(last_starting_line), u'', data_format_c)
                    if iliskili_surec is not None:
                        iliskili_surec_value = cnn.getsinglekoddata('kod_ek_mevzuat_iliskili_surec', 'kod', 'objectid='+str(iliskili_surec))
                        ws.write('C'+str(last_starting_line), iliskili_surec_value.decode('utf-8'), data_format_c)
                    else:
                        ws.write('C'+str(last_starting_line), u'', data_format_c)
                    if veri_paylasimina_etkileri is not None:
                        ws.merge_range('D'+str(last_starting_line)+':E'+str(last_starting_line), veri_paylasimina_etkileri.decode('utf-8'), data_format_c)
                    else:
                        ws.merge_range('D'+str(last_starting_line)+':E'+str(last_starting_line), u'', data_format_c)
                    
                    if etkiledigi_tema_katman is not None:
                        ws.write('F'+str(last_starting_line), etkiledigi_tema_katman.decode('utf-8'), data_format_c)
                    else:
                        ws.write('F'+str(last_starting_line), u'', data_format_c)
                    last_starting_line += 1

            if last_starting_line == 15:
                ws.set_row(14, 15)
                # ws.merge_range('A'+str(last_starting_line)+':F'str(last_starting_line),u'', data_format_c)
                last_starting_line += 1
                ws.set_row(last_starting_line, 42)
            else:
                ws.set_row(last_starting_line-1, 42)
            
            ws.write('A'+str(last_starting_line), u'Coğrafi Veri Paylaşılamama Sebebi', sub_header_format)
            ws.merge_range('B'+str(last_starting_line)+':F'+str(last_starting_line), u'', data_format_r)
            if self.veri_paylasmama_sebep is not None:
                h = 15*len(self.veri_paylasmama_sebep)/100
                ws.write('B'+str(last_starting_line), self.veri_paylasmama_sebep.decode('utf-8'), data_format_r)
                if h < 15:
                    h = 30
                else:
                    h = 40
                ws.set_row(last_starting_line, h)
                
                #ws.write('B'+str(last_starting_line), self.veri_paylasmama_sebep.decode('utf-8'), data_format_r)


            wb.close()
        except Exception as e:
            import sys
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
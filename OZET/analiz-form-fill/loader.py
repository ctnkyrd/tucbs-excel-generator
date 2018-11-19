# -*- coding: utf-8 -*-
import xlsxwriter,sys,os
import psycopg2
import re
from random import randint
from formData import FormData
from pgget import Connection

cnn = Connection()


kurumName = "Hata"




        
# The Sheet Constructor
def sheetConstructor(katmanAdi, sheetName, data):

    try:
        worksheet = workbook.add_worksheet(unicode(sheetName))
        print katmanAdi + "--Started"
        # data loader
        worksheet.write('D2', data.evetHayir(data.katmanDurumu), cell_border)
        if data.katman_aciklama is not None:
            worksheet.write('D1', data.katman_aciklama.decode('utf-8'), cell_border)

        worksheet.write('D4',data.projeksiyon.decode('utf-8')+u' '+data.datum.decode('utf-8'))
        worksheet.write('D3',data.vk_zamansal_gecerlilik.decode('utf-8'),cell_border)
        # worksheet.write('D4', data.stringText(data.veriKoordinat),cell_border)
        worksheet.write('D6', data.stringText(data.veriEksizliklik),cell_border)
        worksheet.write('D5', u'', cell_border)
        worksheet.write('D6', data.stringText(data.veriEksizliklik), cell_border)
        worksheet.write('D7', data.stringText(data.veriMantiksal),cell_border)
        worksheet.write('D8', data.stringText(data.veriKonumsal),cell_border)
        worksheet.write('D9', data.stringText(data.veriZamansal),cell_border)
        worksheet.write('D10', data.stringText(data.veriTematik),cell_border)
        worksheet.write('D11', data.sayisalMi(),cell_border)
        worksheet.write('D12', u'', cell_border)
        worksheet.write('D13', u'', cell_border)
        worksheet.write('D14', data.evetHayir(data.wmsDurum),cell_border)
        worksheet.write('D15', data.wmsWfsStandart(data.wmsDurum),cell_border)
        worksheet.write('D16', u'', cell_border)
        worksheet.write('D17', data.evetHayir(data.wfsDurum),cell_border)
        worksheet.write('D18', data.wmsWfsStandart(data.wfsDurum),cell_border)
        worksheet.write('D19', u'', cell_border)
        worksheet.write('D20', u'', cell_border)
        worksheet.write('D21', data.evetHayir(data.metaveriDurumu),cell_border)
        worksheet.write('D22', u'', cell_border)
        worksheet.write('D23', u'', cell_border)
        worksheet.write('D24', u'', cell_border)
        worksheet.write('D25', u'', cell_border)
        worksheet.write('D26', data.evetHayir(data.metaveriCbsgm),cell_border)
        worksheet.write('D27', u'', cell_border)
        worksheet.write('D28', u'', cell_border)
        worksheet.write('D29', u'', cell_border)
        worksheet.write('D30', u'', cell_border)
        worksheet.write('D31', u'', cell_border)
        worksheet.write('D32', data.cbsBrimi(),cell_border)
        worksheet.write('D33', data.personelDurum(),cell_border)
        worksheet.write('D34', data.donanimDurum(),cell_border)
        worksheet.write('D35', u'', cell_border)
        worksheet.write('D36', u'', cell_border)
        worksheet.write('D37', data.kamunetIpsec(),cell_border)



        # worksheet.write('D1', u'Açıklama',red_header)
        worksheet.write('C2', u'Veri Üretiliyor mu', cell_bold)
        worksheet.write('C3', u'Verinin Güncelliği', cell_bold)
        worksheet.write('C4', u'Verinin Koordinat Bilgisi', cell_bold)
        worksheet.write('C5', u'Veri Kalitesi', cell_bold)
        worksheet.write('C6', u'Verinin Eksizliği (Completeness)', italic_right)
        worksheet.write('C7', u'Mantıksal Tutarlılık (Logical Consistency)', italic_right)
        worksheet.write('C8', u'Konumsal Doğruluk (Positional Accuracy)', italic_right)
        worksheet.write('C9', u'Zamansal Doğruluk (Temporal Accuracy)', italic_right)
        worksheet.write('C10', u'Tematik Doğruluk (Thematic Accuracy)', italic_right)
        worksheet.write('C11', u'Veri Sayısal Formatta mı', cell_bold)
        worksheet.write('C12', u'Verinin öznitelik bilgileri var mı', cell_bold)
        worksheet.write('C13', u'Verinin öznitelik bilgileri standart mı', cell_bold)
        worksheet.write('C14', u'WMS var mı', cell_bold)
        worksheet.write('C15', u'WMS standardı', cell_bold)
        worksheet.write('C16', u'WMS içerik tam geliyor mu?', cell_bold)
        worksheet.write('C17', u'WFS var mı', cell_bold)
        worksheet.write('C18', u'WFS standardı', cell_bold)
        worksheet.write('C19', u'WMS içerik tam geliyor mu?', cell_bold)
        worksheet.write('C20', u'Anlık veri servis ediliyor mu?', cell_bold)
        worksheet.write('C21', u'Metaverisi var mı', cell_bold)
        worksheet.write('C22', u'Metaveri içerikleri tam mı', cell_bold)
        worksheet.write('C23', u'Metaveri tutulma formatı', cell_bold)
        worksheet.write('C24', u'Metaveri servisi var mı (harvest)', cell_bold)
        worksheet.write('C25', u'Metaveri güncelliği', cell_bold)
        worksheet.write('C26', u'Geoportale aktarılma durumu', cell_bold)
        worksheet.write('C27', u'Geoportale aktarılma yöntemi', cell_bold)
        worksheet.write('C28', u'Geoportale aktarılma sırasında içerikleri tam mı', cell_bold)
        worksheet.write('C29', u'Metaveri içinde WMS var mı', cell_bold)
        worksheet.write('C30', u'Metaveri içinde WFS var mı', cell_bold)
        worksheet.write('C31', u'Metaverinin kurum ve kuruluşlara erişim durumu', cell_bold)
        worksheet.write('C32', u'CBS Birimi Var mı', cell_bold)
        worksheet.write('C33', u'CBS Personeli Var mı/Yeterli mi', cell_bold)
        worksheet.write('C34', u'Donanım var mı/Yeterli mi', cell_bold)
        worksheet.write('C35', u'Yazılım Varmı/Yeterli mi', cell_bold)
        worksheet.write('C36', u'Sunucu Var mı/Yeterli mi', cell_bold)
        worksheet.write('C37', u'Kamunet/IPSECVPN\'e  bağlı mı', cell_bold)
        worksheet.set_column('C:C', 45)
        worksheet.set_column('D:D', 50)

        #B:B column
        worksheet.write('B2', u'1',text_red)
        worksheet.write('B3', u'2',text_red)
        worksheet.write('B4', u'3',text_red)
        worksheet.write('B5', u'4',text_red)
        worksheet.write('B6', u'4.1',text_red)
        worksheet.write('B7', u'4.2',text_red)
        worksheet.write('B8', u'4.3',text_red)
        worksheet.write('B9', u'4.4',text_red)
        worksheet.write('B10', u'4.5',text_red)
        worksheet.write('B11', u'5',text_red)
        worksheet.write('B12', u'6',text_red)
        worksheet.write('B13', u'7',text_red)
        worksheet.write('B14', u'8',text_green)
        worksheet.write('B15', u'',text_green)
        worksheet.write('B16', u'9',text_green)
        worksheet.write('B17', u'10',text_green)
        worksheet.write('B18', u'',text_green)
        worksheet.write('B19', u'11',text_green)
        worksheet.write('B20', u'12',text_green)
        worksheet.write('B21', u'13',text_yellow)
        worksheet.write('B22', u'14',text_yellow)
        worksheet.write('B23', u'15',text_yellow)
        worksheet.write('B24', u'16',text_yellow)
        worksheet.write('B25', u'17',text_yellow)
        worksheet.write('B26', u'18',text_yellow)
        worksheet.write('B27', u'19',text_yellow)
        worksheet.write('B28', u'20',text_yellow)
        worksheet.write('B29', u'21',text_yellow)
        worksheet.write('B30', u'22',text_yellow)
        worksheet.write('B31', u'23',text_yellow)
        worksheet.write('B32', u'24',text_orange)
        worksheet.write('B33', u'25',text_orange)
        worksheet.write('B34', u'26',text_orange)
        worksheet.write('B35', u'27',text_orange)
        worksheet.write('B36', u'28',text_orange)
        worksheet.write('B37', u'29',text_orange)

        #the merges
        worksheet.merge_range('A1:C1', unicode(katmanAdi),red_header)
        worksheet.merge_range('A2:A13', u'COĞRAFİ VERİ',rotated_text_red)
        worksheet.merge_range('A14:A20', u'COĞRAFİ VERİ SERVİSİ',rotated_text_green)
        worksheet.merge_range('A21:A31', u'METAVERİ',rotated_text_yellow)
        worksheet.merge_range('A32:A37', u'DONANIM,YAZILIM ve PERSONEL',rotated_text_orange)
    except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)

def checkNames(katmanListesi,katmanAdi):
    if(katmanAdi in katmanListesi):
        katmanAdi = katmanAdi + "_"+str(randint(0, 10000))
        return katmanAdi
    else:
        return katmanAdi

def correctSheetName(sheetName):
    sheetName = ''.join(e for e in sheetName if e.isalnum())
    if (len(sheetName)>=25):
        sheetName = sheetName[:24]
        return sheetName
    else:
        return sheetName

try:


    kurumListe = cnn.getlistofdata('kurum','objectid, adi','analiz_tamamlandi_first is true')

    for i in kurumListe:
        print i[0], "-----" , i[1].decode('utf-8')

    user_input = raw_input("Kurum ID'sini giriniz(Tum Kurumlar icin -1): ")
    conn = psycopg2.connect("dbname='tucbsdata' user='postgres' host='192.168.30.136' password='Ankara123'")
    cur = conn.cursor()

    if int(user_input) == -1:
        cur.execute("select objectid, adi from kurum where analiz_tamamlandi_first = true")
    else:
        cur.execute("select objectid, adi from kurum where analiz_tamamlandi_first = true and objectid = "+ str(user_input))
    
    allKurum = cur.fetchall()
    for kurum in allKurum:
        kurumName = kurum[1].decode('utf-8')
        folderLocation = "/Forms"
        documentName = kurumName

        workbook = xlsxwriter.Workbook('Forms/'+kurumName.rstrip()+'.xlsx')
        kurumId = kurum[0]
        cur_ek2 = conn.cursor()
        cur_ek2.execute("""select objectid from ek_2_cografi_veri_analizi where kurum = %s and geodurum is true""", [kurumId])
        ek_2 = cur_ek2.fetchone()[0]
        cur_xkatman = conn.cursor()
        cur_xkatman.execute("""select * from x_ek_2_tucbs_veri_katmani where ek_2 = %s and geodurum is true""", [ek_2])
        allXKatman = cur_xkatman.fetchall()

        cur_per = conn.cursor()
        cur_per.execute(""" select cbs_birimi_var,  personel_yeterli from ek_cbs_organizasyon_birimleri_analizi where kurum = %s and geodurum is true""", [kurumId])
        personel = cur_per.fetchall()[0]
        cbsBirimi = personel[0]
        cbsPersonel = personel[1]
       
        cur_don = conn.cursor()
        cur_don.execute(""" select sunucu_yeterli, kamunet_agina_bagli, ipsecvpn_uygun from ek_donanim_yazilim_ag_guvenlik where kurum = %s and geodurum is true""", [kurumId])
        donanim = cur_don.fetchall()[0]
        sunucuDurum = donanim[0]
        kamunetAgi = donanim[1]
        ipsecVpn = donanim[2]

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('Forms/'+kurumName+'.xlsx')
        #cell formats
        cell_bold = workbook.add_format()
        cell_bold.set_bold()
        cell_bold.set_border()

        cell_border = workbook.add_format()
        cell_border.set_border()
        cell_border.set_text_wrap()

        italic_right = workbook.add_format()
        italic_right.set_italic()
        italic_right.set_align('right')
        italic_right.set_border()

        rotated_text_red = workbook.add_format()
        rotated_text_red.set_rotation(90)
        rotated_text_red.set_align('center')
        rotated_text_red.set_align('vcenter')
        rotated_text_red.set_text_wrap()
        rotated_text_red.set_bg_color('#ffc7ce')
        rotated_text_red.set_border()
        rotated_text_red.set_color('#9c0006')

        rotated_text_green = workbook.add_format()
        rotated_text_green.set_rotation(90)
        rotated_text_green.set_align('center')
        rotated_text_green.set_align('vcenter')
        rotated_text_green.set_text_wrap()
        rotated_text_green.set_bg_color('#c6efce')
        rotated_text_green.set_border()
        rotated_text_green.set_color('#006100')

        rotated_text_yellow = workbook.add_format()
        rotated_text_yellow.set_rotation(90)
        rotated_text_yellow.set_align('center')
        rotated_text_yellow.set_align('vcenter')
        rotated_text_yellow.set_text_wrap()
        rotated_text_yellow.set_bg_color('#ffeb9c')
        rotated_text_yellow.set_border()
        rotated_text_yellow.set_color('#9c6500')

        rotated_text_orange = workbook.add_format()
        rotated_text_orange.set_rotation(90)
        rotated_text_orange.set_align('center')
        rotated_text_orange.set_align('vcenter')
        rotated_text_orange.set_text_wrap()
        rotated_text_orange.set_bg_color('#ffcc99')
        rotated_text_orange.set_border()
        rotated_text_orange.set_color('#3f3f76')

        text_red = workbook.add_format()
        text_red.set_border()
        text_red.set_bg_color('#ffc7ce')
        text_red.set_color('#9c0006')

        text_green = workbook.add_format()
        text_green.set_border()
        text_green.set_bg_color('#c6efce')
        text_green.set_color('#006100')

        text_yellow = workbook.add_format()
        text_yellow.set_border()
        text_yellow.set_bg_color('#ffeb9c')
        text_yellow.set_color('#9c6500')

        text_orange = workbook.add_format()
        text_orange.set_border()
        text_orange.set_bg_color('#ffcc99')
        text_orange.set_color('#3f3f76')

        red_header = workbook.add_format()
        red_header.set_align('center')
        red_header.set_color('red')
        red_header.set_border()
        
        katmanNames = []
        # the code goes here
        for en, katman in enumerate(allXKatman):
            katmanAdi = unicode(katman[2].decode('utf-8'))
            # katmanAdi = checkNames(katmanNames, katmanAdi)
            # katmanNames.append(katmanAdi)
            # a = FormData('2','a','Koordinat','Eksiksizlik','mantiksal','konumsal', 'zamansal', 'tematik')
            sheetName = unicode(str(en)) + u"_" +correctSheetName(katmanAdi)
            data = FormData(katman[3], katman[21], katman[12], katman[23], katman[27], 
                            katman[28],katman[29],katman[30],katman[11], katman[11],
                            katman[43], katman[18], katman[44], katman[18],
                            katman[42], katman[41],
                            cbsBirimi, cbsPersonel, sunucuDurum, kamunetAgi, ipsecVpn,
                            katman[61], katman[0], katman[62], katman[63], katman[80]
                            )
            sheetConstructor(katmanAdi,  sheetName, data)
            
        workbook.close()
except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            print(exc_type, fname, exc_tb.tb_lineno)
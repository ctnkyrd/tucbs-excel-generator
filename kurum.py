# -*- coding: utf-8 -*-
from pgget import Connection
cnn = Connection()
class Kurum:
    def __init__(self, oid):
        self.oid = oid
        self.kurum_kodu = cnn.getSingledataByOid('kurum', 'kurum_kodu',self.oid)
        self.ust_kodu = cnn.getSingledataByOid('kurum', 'ust_kodu',self.oid)
        # Coğrafi veri analiz formu adı
        self.adi = cnn.getSingledataByOid('kurum', 'adi',self.oid).decode('utf-8')
        # 1 kurum, 2 belediye
        self.tipi = cnn.getSingledataByOid('kurum', 'tipi',self.oid)
        if self.tipi == 1:
            self.tip = u'Kurum'
            self.bakanlik = cnn.getsinglekoddata('kurum', 'adi', 'kurum_kodu='+str(self.ust_kodu)).decode('utf-8')
        elif self.tipi == 2:
            self.tip = u'Belediye'
            self.bakanlik = cnn.getsinglekoddata('kurum', 'adi', 'kurum_kodu='+str(self.ust_kodu)).decode('utf-8')
        else:
            self.tip = u'Belirlenemedi'
            self.bakanlik = u'Belirlenemedi'
        # get ek2 information objectid and birim
        self.ek2 = cnn.getRowOfData('ek_2_cografi_veri_analizi','objectid, birim','geodurum is true')
        self.ek2_oid = self.ek2[0]
        self.birim = self.ek2[1].decode('utf-8')
    
    def getOid(self):
        return self.oid
    
    class VeriKatmani():
        def __init__(self,katman):
            self.katman = katman
            self.oid = cnn.executeSql("select * from x_ek_2_tucbs_veri_katmani where ek_2 = "+str(self.katman.oid))
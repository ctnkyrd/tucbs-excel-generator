# -*- coding: utf-8 -*-
from pgget import Connection
cnn = Connection()

class Kurum:
    def __init__(self, oid, verikatmani=None):
        self.oid = oid
        self.kurum_kodu = cnn.getSingledataByOid('kurum', 'kurum_kodu',self.oid)
        self.ust_kodu = cnn.getSingledataByOid('kurum', 'ust_kodu',self.oid)
        self.k_adi = cnn.getSingledataByOid('kurum', 'k_adi', self.oid)
        # Coğrafi veri analiz formu adı
        self.adi = cnn.getSingledataByOid('kurum', 'adi',self.oid).decode('utf-8')
        # katman listesi
        if verikatmani is None:
            self.verikatmani = []
        else:
            self.verikatmani = verikatmani
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
        self.ek2 = cnn.getRowOfData('ek_2_cografi_veri_analizi','objectid, birim','geodurum is true and kurum='+str(self.oid))
        self.ek2_oid = self.ek2[0]
        if self.ek2[1] is not None:
            self.birim = self.ek2[1].decode('utf-8')
        else:
            self.birim = None

    def add_veri_katmani(self, katman):
        if not katman in self.verikatmani:
            self.verikatmani.append(katman)
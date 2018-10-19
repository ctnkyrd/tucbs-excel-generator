# -*- coding: utf-8 -*-
from pgget import Connection
cnn = Connection()
class Kurum:
    def __init__(self, oid, adi, tipi, kurum_kodu, ustkod):
        self.oid = oid
        self.kurum_kodu = kurum_kodu
        # Coğrafi veri analiz formu adı
        self.adi = adi
        # 1 kurum, 2 belediye
        if tipi == 1:
            self.tipi = 'Kurum'
            self.bakanlik = cnn.getsinglekoddata('kurum', 'adi', 'kurum_kodu='+str(ustkod)).decode('utf-8')
        elif tipi == 2:
            self.tipi = 'Belediye'
            self.bakanlik = cnn.getsinglekoddata('kurum', 'adi', 'kurum_kodu='+str(ustkod)).decode('utf-8')
        else:
            self.tipi = 'Belirlenemedi'
            self.bakanlik = 'Belirlenemedi'
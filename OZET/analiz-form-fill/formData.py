from pgget import Connection

cnn = Connection()

class FormData: 
    def __init__(self, katmanDurumu, veriGuncelligi, veriKoordinat, 
                veriEksizliklik, veriMantiksal, veriKonumsal, veriZamansal, veriTematik,
                veriSayisal, veriOznitelik, wmsDurum, wmsOgc, wfsDurum, wfsOgc,
                metaveriDurumu, metaveriCbsgm,
                cbsBirimiDurumu, cbsPersonelDurumu, donanimYeterli, kamunetBagli, ipsecBagli,
                katman_aciklama, oid, projeksiyon, datum, vk_zamansal_gecerlilik
                ):
        self.oid = oid
        self.katmanDurumu = katmanDurumu
        self.veriGuncelligi = veriGuncelligi
        self.veriKoordinat = veriKoordinat

        self.veriEksizliklik = veriEksizliklik
        self.veriMantiksal = veriMantiksal
        self.veriKonumsal = veriKonumsal
        self.veriZamansal = veriZamansal
        self.veriTematik = veriTematik

        self.wmsDurum = wmsDurum
        self.wmsOgc = wmsOgc
        self.wfsDurum = wfsDurum
        self.wfsOgc = wfsOgc

        self.metaveriDurumu = metaveriDurumu
        self.metaveriCbsgm = metaveriCbsgm

        self.cbsBirimiDurumu = cbsBirimiDurumu
        self.cbsPersonelDurumu = cbsPersonelDurumu
        self.donanimYeterli = donanimYeterli
        self.kamunetBagli = kamunetBagli
        self.ipsecBagli = ipsecBagli

        self.veriSayisal = veriSayisal
        self.veriOznitelik = veriOznitelik

        self.katman_aciklama = katman_aciklama

        if vk_zamansal_gecerlilik == 1:
            self.vk_zamansal_gecerlilik = u'Guncel'
        else:
            self.vk_zamansal_gecerlilik = u'Guncel Degil'


        if projeksiyon is not None:
            self.projeksiyon = cnn.getsinglekoddata('kod_ek_2_projeksiyon', 'kod', 'objectid='+str(projeksiyon))
        else:
            self.projeksiyon = ''
        
        if datum is not None:
            self.datum = cnn.getsinglekoddata('kod_ek_2_datum', 'kod', 'objectid='+str(datum))
        else:
            self.datum =''

        

    def evetHayir(self, field):
        if field is True:
            return u'Evet'
        else:
            return u'Hayir'
    def veriGuncellik(self):
        if self.veriGuncelligi is not None:
            return u'Guncel'
        else:
            return u'Guncel Degil'
    def stringText(self, string):
        if string is None:
            return " "
        else:
            return string.decode('utf-8')
    def sayisalMi(self):
        if self.veriSayisal > 2:
            return u'Evet'
        elif self.veriSayisal <= 2:
            return u'Hayir'
        else:
            return u' '
    def wmsWfsStandart(self, servis):
        if self.wmsOgc is True and servis is True:
            return u'OGC'
        elif servis is False:
            return u''
        else: 
            return u'Stardart Degil'
    def cbsBrimi(self):
        if self.cbsBirimiDurumu is True:
            return u'Var'
        else:
            return u'Yok'
    def personelDurum(self):
        if self.cbsPersonelDurumu is True:
            return u'Var - Yeterli'
        else:
            return u'Yetersiz'
    def donanimDurum(self):
        if self.donanimYeterli is True:
            return u'Var - Yeterli'
        else:
            return u'Yetersiz'
    def kamunetIpsec(self):
        if self.kamunetBagli or self.ipsecBagli is True:
            return u'Evet'
        else:
            return u'False'
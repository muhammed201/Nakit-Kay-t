Attribute VB_Name = "SayiOku"
Global okunacak, okunacakyeni, okunacakyeni_real

Sub Virgulden_Once(okunacak)
Dim birler, onlar, diger
Dim sayi As Double
Dim ucluk, kalan As Integer
Dim l, m, r As Byte
birler = Array("", "Bir", "Ýki", "Üç", "Dört", "Beþ", "Altý", "Yedi", "Sekiz", "Dokuz")
    onlar = Array("", "On", "Yirmi", "Otuz", "Kýrk", "Elli", "Altmýþ", "Yetmiþ", "Seksen", "Doksan")
    diger = Array("", "Bin", "Milyon", "Milyar", "Trilyon")
        sayi = Val(okunacak)
    okunacakyeni = ""
    ucluk = Int((Len(okunacak) - 1) / 3)
    For i = ucluk To 0 Step -1
        kalan = Int(sayi / 1000 ^ i)
        sayi = sayi - kalan * 1000 ^ i
        If kalan = 0 Then GoTo Son
        l = Int(kalan / 100)
        m = Int((kalan - l * 100) / 10)
        r = Int(kalan - l * 100 - m * 10)
        If l > 1 Then okunacakyeni = okunacakyeni + birler(l) 'BirYüz olmamasý için
        If l > 0 Then okunacakyeni = okunacakyeni & "Yüz"
        If i = 1 And kalan = 1 Then okunacakyeni = okunacakyeni & diger(i): GoTo Son 'BirBin olmamasý için
        okunacakyeni = okunacakyeni & onlar(m) & birler(r) & diger(i)
Son:
    Next
End Sub

Sub Virgulden_Sonra(okunacak)
Dim birler, onlar, diger
Dim sayi As Double
Dim ucluk, kalan As Integer
Dim l, m, r As Byte
birler = Array("", "Bir", "Ýki", "Üç", "Dört", "Beþ", "Altý", "Yedi", "Sekiz", "Dokuz")
    onlar = Array("", "On", "Yirmi", "Otuz", "Kýrk", "Elli", "Altmýþ", "Yetmiþ", "Seksen", "Doksan")
    diger = Array("", "Bin", "Milyon", "Milyar", "Trilyon")
        sayi = Val(okunacak)
    okunacakyeni_real = ""
    ucluk = Int((Len(okunacak) - 1) / 3)
    For i = ucluk To 0 Step -1
        kalan = Int(sayi / 1000 ^ i)
        sayi = sayi - kalan * 1000 ^ i
        If kalan = 0 Then GoTo Son
        l = Int(kalan / 100)
        m = Int((kalan - l * 100) / 10)
        r = Int(kalan - l * 100 - m * 10)
        If l > 1 Then okunacakyeni_real = okunacakyeni_real + birler(l) 'BirYüz olmamasý için
        If l > 0 Then okunacakyeni_real = okunacakyeni_real & "Yüz"
        If i = 1 And kalan = 1 Then okunacakyeni_real = okunacakyeni_real & diger(i): GoTo Son 'BirBin olmamasý için
        okunacakyeni_real = okunacakyeni_real & onlar(m) & birler(r) & diger(i)
Son:
    Next
End Sub

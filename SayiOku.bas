Attribute VB_Name = "SayiOku"
Global okunacak, okunacakyeni, okunacakyeni_real

Sub Virgulden_Once(okunacak)
Dim birler, onlar, diger
Dim sayi As Double
Dim ucluk, kalan As Integer
Dim l, m, r As Byte
birler = Array("", "Bir", "�ki", "��", "D�rt", "Be�", "Alt�", "Yedi", "Sekiz", "Dokuz")
    onlar = Array("", "On", "Yirmi", "Otuz", "K�rk", "Elli", "Altm��", "Yetmi�", "Seksen", "Doksan")
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
        If l > 1 Then okunacakyeni = okunacakyeni + birler(l) 'BirY�z olmamas� i�in
        If l > 0 Then okunacakyeni = okunacakyeni & "Y�z"
        If i = 1 And kalan = 1 Then okunacakyeni = okunacakyeni & diger(i): GoTo Son 'BirBin olmamas� i�in
        okunacakyeni = okunacakyeni & onlar(m) & birler(r) & diger(i)
Son:
    Next
End Sub

Sub Virgulden_Sonra(okunacak)
Dim birler, onlar, diger
Dim sayi As Double
Dim ucluk, kalan As Integer
Dim l, m, r As Byte
birler = Array("", "Bir", "�ki", "��", "D�rt", "Be�", "Alt�", "Yedi", "Sekiz", "Dokuz")
    onlar = Array("", "On", "Yirmi", "Otuz", "K�rk", "Elli", "Altm��", "Yetmi�", "Seksen", "Doksan")
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
        If l > 1 Then okunacakyeni_real = okunacakyeni_real + birler(l) 'BirY�z olmamas� i�in
        If l > 0 Then okunacakyeni_real = okunacakyeni_real & "Y�z"
        If i = 1 And kalan = 1 Then okunacakyeni_real = okunacakyeni_real & diger(i): GoTo Son 'BirBin olmamas� i�in
        okunacakyeni_real = okunacakyeni_real & onlar(m) & birler(r) & diger(i)
Son:
    Next
End Sub

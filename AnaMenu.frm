VERSION 5.00
Begin VB.Form AnaMenu 
   Caption         =   "Nakit Kayýt ANA MENÜ"
   ClientHeight    =   7470
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   12045
   Icon            =   "AnaMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Düzenlenecek Nakit Tipini Seçiniz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   5760
      Width           =   11775
      Begin VB.OptionButton Option16 
         Caption         =   "Öðr. Sgr. Prim"
         Height          =   495
         Left            =   1920
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Huzur Hak. Öd."
         Height          =   495
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Yolluk(Tedavi)"
         Height          =   495
         Left            =   8280
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Yolluk(Sürekli)"
         Height          =   495
         Left            =   6720
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Düzenle"
         Height          =   615
         Left            =   9960
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Eczane"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Hastane"
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Poliklinik"
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Yolluk(Geçici)"
         Height          =   495
         Left            =   5160
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menü"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Frame Frame5 
         Caption         =   "Eklenecek Tipi Seçiniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5415
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton Command12 
            Caption         =   "Öðrenci Sigorta Primi Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   4680
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Huzur Hakký Ödemeleri Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   4080
            Width           =   1575
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Yolluk(Tedavi) Nakit Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   33
            Top             =   3480
            Width           =   1575
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Yolluk(Sürekli) Nakit Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   2880
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Yolluk(Geçici) Nakit Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Poliklinik Nakit Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Hastane Nakit Kaydý"
            Height          =   435
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Eczane Nakit Kaydý"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1575
         End
         Begin VB.Timer trmUnload 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1680
            Top             =   4440
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listelenecek Nakit Tiplerini Seçiniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   2040
         TabIndex        =   7
         Top             =   3120
         Width           =   9615
         Begin VB.CheckBox Check8 
            Caption         =   "Öðr. Sgr. Prim Listesi"
            Height          =   375
            Left            =   2400
            TabIndex        =   38
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Huzur Hak. Öd Listesi"
            Height          =   615
            Left            =   360
            TabIndex        =   37
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Yolluk Nakit (Tedavi) Listesi"
            Height          =   375
            Left            =   4320
            TabIndex        =   27
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Yolluk Nakit (Sürekli) Listesi"
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Listele"
            Height          =   855
            Left            =   8160
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Yolluk Nakit(Geçici) Listesi"
            Height          =   375
            Left            =   2400
            TabIndex        =   11
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Poliklinik Nakit Listesi"
            Height          =   375
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Hastane Nakit Listesi"
            Height          =   375
            Left            =   2400
            TabIndex        =   9
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Eczane Nakit Listesi"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Timer trmLoad 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   4680
      End
      Begin VB.Frame Frame2 
         Caption         =   "Yazdýrýlacak Nakit Tipini Seçiniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2775
         Left            =   2040
         TabIndex        =   1
         Top             =   360
         Width           =   9615
         Begin VB.OptionButton Option14 
            Caption         =   "Huzur Hak.Öd."
            Height          =   495
            Left            =   4680
            TabIndex        =   36
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Öðr. Sgr. Prim"
            Height          =   495
            Left            =   4680
            TabIndex        =   35
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Yolluk(Tedavi)"
            Height          =   495
            Left            =   3240
            TabIndex        =   32
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Yolluk(Sürekli)"
            Height          =   495
            Left            =   1440
            TabIndex        =   31
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Yolluk(Geçici)"
            Height          =   495
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Sayfayý Yenile"
            Height          =   375
            Left            =   8040
            TabIndex        =   24
            Top             =   120
            Width           =   1455
         End
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   3480
            Top             =   2160
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Yazdýr"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8040
            TabIndex        =   6
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   2040
            Width           =   3135
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Poliklinik"
            Height          =   495
            Left            =   3240
            TabIndex        =   4
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Hastane"
            Height          =   495
            Left            =   1440
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Eczane"
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menü"
      Begin VB.Menu cikis 
         Caption         =   "Çýkýþ"
      End
   End
   Begin VB.Menu yardim 
      Caption         =   "Yardým"
      Begin VB.Menu hakkinda 
         Caption         =   "Hakkýnda"
      End
   End
End
Attribute VB_Name = "AnaMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public yazdirilacak As String




Private Sub cikis_Click()
Unload Me
End Sub

Private Sub Combo1_Change()
Call yetki
End Sub

Private Sub Combo1_Click()
Call yetki
End Sub

Private Sub Command1_Click()
Eczane.Show
Unload Hastane
Unload Poliklinik
Unload Yolluk
Unload Me
Unload Yolluksurekli
Unload Yolluktedavi
Unload ogrprim
End Sub

Private Sub Command10_Click()
Yolluktedavi.Show
Unload Yolluksurekli
Unload Hastane
Unload Eczane
Unload Poliklinik
Unload Yolluk
Unload Me
Unload ogrprim
End Sub

Private Sub Command11_Click()
huzurhakki.Show
Unload Yolluktedavi
Unload Yolluksurekli
Unload Hastane
Unload Eczane
Unload Poliklinik
Unload Yolluk
Unload Me
Unload ogrprim
End Sub

Private Sub Command12_Click()
ogrprim.Show
Unload huzurhakki
Unload Yolluktedavi
Unload Yolluksurekli
Unload Hastane
Unload Eczane
Unload Poliklinik
Unload Yolluk
Unload Me
End Sub

Private Sub Command2_Click()
Hastane.Show
Unload Eczane
Unload Poliklinik
Unload Yolluk
Unload Yolluksurekli
Unload Yolluktedavi
Unload Me
Unload huzurhakki
Unload ogrprim
End Sub

Private Sub Command3_Click()
Poliklinik.Show
Unload Hastane
Unload Eczane
Unload Yolluk
Unload Yolluksurekli
Unload Yolluktedavi
Unload Me
Unload huzurhakki
Unload ogrprim
End Sub

Private Sub Command4_Click()
Yolluk.Show
Unload Hastane
Unload Eczane
Unload Poliklinik
Unload Yolluksurekli
Unload Yolluktedavi
Unload Me
Unload huzurhakki
Unload ogrprim
End Sub

Private Sub Command5_Click()
Call kontrol
End Sub

Private Sub Command6_Click()
If Check1.Value = 1 Then
Eczaneliste.Show
End If

If Check2.Value = 1 Then
Hastaneliste.Show
End If

If Check3.Value = 1 Then
Poliklinikliste.Show
End If

If Check4.Value = 1 Then
Yollukliste.Show
End If

If Check5.Value = 1 Then
Yolluksurekliliste.Show
End If

If Check6.Value = 1 Then
Yolluktedavililiste.Show
End If

If Check7.Value = 1 Then
huzurhakkiliste.Show
End If

If Check8.Value = 1 Then
ogrprimliste.Show
End If

End Sub

Private Sub Command7_Click()
If Option8.Value = True Then
eczane_duzelt.Show
ElseIf Option7.Value = True Then
Hastane_duzelt.Show
ElseIf Option6.Value = True Then
Poliklinik_duzelt.Show
ElseIf Option5.Value = True Then
Yolluk_duzelt.Show
ElseIf Option11.Value = True Then
Yolluksurekli_duzelt.Show
ElseIf Option12.Value = True Then
Yolluktedavi_duzelt.Show
ElseIf Option15.Value = True Then
huzurhakki_duzelt.Show
ElseIf Option16.Value = True Then
ogrprim_duzelt.Show
End If
End Sub

Private Sub Command8_Click()
Call veri_ac(False, False)
Call tablo_ac("Select * from Eczane")
Call listele
yazdirilacak = "Eczane"
Option1.Value = True
End Sub

Private Sub Command9_Click()
Yolluksurekli.Show
Unload Hastane
Unload Eczane
Unload Poliklinik
Unload Yolluk
End Sub

Private Sub Form_Load()
SetWindowRgn hwnd, Rgn1, True
SetTransparent hwnd, 0
trmLoad.Enabled = True

Me.Left = Screen.Width / 10
Me.Top = Screen.Height / 10
Call veri_ac(False, False)
Call tablo_ac("Select * from Eczane")
Call listele
yazdirilacak = "Eczane"
End Sub


Private Sub Form_Unload(Cancel As Integer)
trmUnload = True
trmLoad = False
End Sub

Private Sub hakkinda_Click()
frmAbout.Show
End Sub

Private Sub Option1_Click()
Call tablo_ac("Select * from Eczane")
Call listele
Call yetki
yazdirilacak = "Eczane"
End Sub

Private Sub Option10_Click()
Call tablo_ac("Select * from Yolluktedavi")
Call listele
Call yetki
yazdirilacak = "Yolluktedavi"
End Sub



Private Sub Option13_Click()
Call tablo_ac("Select * from ogrprim")
Call listele
Call yetki
yazdirilacak = "ogrprim"
End Sub

Private Sub Option14_Click()
Call tablo_ac("Select * from huzurhakki")
Call listele
Call yetki
yazdirilacak = "huzurhakki"
End Sub

Private Sub Option2_Click()
Call tablo_ac("Select * from Hastane")
Call listele
Call yetki
yazdirilacak = "Hastane"
End Sub

Private Sub Option3_Click()
Call tablo_ac("Select * from Poliklinik")
Call listele
Call yetki
yazdirilacak = "Poliklinik"
End Sub


Sub listele()
Combo1.Clear
Do While Not tablo.EOF
Combo1.AddItem tablo!tcno & " - " & tablo!adi
tablo.MoveNext
Loop
End Sub
Sub yetki()
If Combo1.Text = Empty Then
Command5.Enabled = False
Else
Command5.Enabled = True
End If
End Sub
Sub kontrol()
On Local Error GoTo hata
Dim bul As Byte
bul = InStr(Combo1.Text, "-") - 2

Call tablo_ac("Select * from " & yazdirilacak & " where tcno='" & Trim(Mid(Combo1.Text, 1, bul)) & "'")
If tablo.RecordCount = 0 Then
MsgBox "Böyle bir kayýt yok", vbExclamation
Exit Sub
Else
Call yazdir
End If

hata:
If Err = 6 Then
MsgBox "Yanlýþ formatta bir deðer girdiniz", vbCritical
End If

End Sub

Sub yazdir()
'*************************BURAYI UNUTMA***************
Dim bul, tamsayi_kismi, real_kismi
'*************************BURAYI UNUTMA***************


Dim Ex As Excel.Application
Set Ex = New Excel.Application
   Ex.Visible = True   'Görünmesini istemiyorsak False
   'Ex.Workbooks.Add 'Yeni bir çalýþma sayfasý oluþturmak için kullanýrýz.
   'Ex.Workbooks.Open ("D:\Muhammed\ReSearch\nakit1.xls") 'Bunu da eðer hazýr bir excel sayfamýz var da onun üzerinde çalýþmak istiyorsak Workbooks.Add yerine kullanabiliriz.
   Ex.Workbooks.Open (App.Path & "\nakit1.xls")
   
   If yazdirilacak = "Eczane" Then
    Ex.Sheets("Eczane").Range("T3").Value = tablo!adi
    Ex.Sheets("Eczane").Range("T4").Value = tablo!tcno
    Ex.Sheets("Eczane").Range("T5").Value = tablo!banka
    Ex.Sheets("Eczane").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Eczane").Range("T7").Value = tablo!vd
    Ex.Sheets("Eczane").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Eczane").Range("C46").Value - Ex.Sheets("Eczane").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Eczane").Range("C46").Value - Ex.Sheets("Eczane").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Eczane").Range("C46").Value - Ex.Sheets("Eczane").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Eczane").Range("C46").Value - Ex.Sheets("Eczane").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Eczane").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Eczane").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Eczane").Range("L46").Value = Round(Ex.Sheets("Eczane").Range("C46").Value - Ex.Sheets("Eczane").Range("I46").Value, 2)
    
           
        
    
    bul = InStr(Ex.Sheets("Eczane").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Eczane").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Eczane").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Eczane").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Eczane").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Eczane").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("Eczane").Range("C46").Value = tamsayi_kismi & "," & real_kismi
    
    
    
    
    
        
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    
    Ex.Sheets("Eczane").Range("R50").Value = tablo!fatura
    Ex.Sheets("Eczane").Range("R51").Value = tablo!sevk
    Ex.Sheets("Eczane").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Eczane").Range("C46").Value = "=P10"
    Ex.Sheets("Eczane").Range("L46").Value = "=P10-P15"
    
    'Ex.Sheets("Poliklinik").PrintOut 'Yazdýr
   ElseIf yazdirilacak = "Hastane" Then
    Ex.Sheets("Hastane").Range("T3").Value = tablo!adi
    Ex.Sheets("Hastane").Range("T4").Value = tablo!tcno
    Ex.Sheets("Hastane").Range("T5").Value = tablo!banka
    Ex.Sheets("Hastane").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Hastane").Range("T7").Value = tablo!vd
    Ex.Sheets("Hastane").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Hastane").Range("C46").Value - Ex.Sheets("Hastane").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Hastane").Range("C46").Value - Ex.Sheets("Hastane").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Hastane").Range("C46").Value - Ex.Sheets("Hastane").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Hastane").Range("C46").Value - Ex.Sheets("Hastane").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Hastane").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Hastane").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Hastane").Range("L46").Value = Round(Ex.Sheets("Hastane").Range("C46").Value - Ex.Sheets("Hastane").Range("I46").Value, 2)
        
    bul = InStr(Ex.Sheets("Hastane").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Hastane").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Hastane").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Hastane").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Hastane").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Hastane").Range("L44").Value = okunacakyeni_real
        
    Ex.Sheets("Hastane").Range("C46").Value = tamsayi_kismi & "," & real_kismi
'****************Bu satýrý  eklemeyi unutma*********************************************

   
    
    Ex.Sheets("Hastane").Range("R50").Value = tablo!fatura
    Ex.Sheets("Hastane").Range("R51").Value = tablo!sevk
    Ex.Sheets("Hastane").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Hastane").Range("C46").Value = "=P10"
    Ex.Sheets("Hastane").Range("L46").Value = "=P10-P15"
    
    'Ex.Sheets("Poliklinik").PrintOut 'Yazdýr
   ElseIf yazdirilacak = "Poliklinik" Then
    Ex.Sheets("Poliklinik").Range("T3").Value = tablo!adi
    Ex.Sheets("Poliklinik").Range("T4").Value = tablo!tcno
    Ex.Sheets("Poliklinik").Range("T5").Value = tablo!banka
    Ex.Sheets("Poliklinik").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Poliklinik").Range("T7").Value = tablo!vd
    Ex.Sheets("Poliklinik").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Poliklinik").Range("C46").Value - Ex.Sheets("Poliklinik").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Poliklinik").Range("C46").Value - Ex.Sheets("Poliklinik").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Poliklinik").Range("C46").Value - Ex.Sheets("Poliklinik").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Poliklinik").Range("C46").Value - Ex.Sheets("Poliklinik").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Poliklinik").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Poliklinik").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Poliklinik").Range("L46").Value = Round(Ex.Sheets("Poliklinik").Range("C46").Value - Ex.Sheets("Poliklinik").Range("I46").Value, 2)
    
    
    
    
    bul = InStr(Ex.Sheets("Poliklinik").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Poliklinik").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Poliklinik").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Poliklinik").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Poliklinik").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Poliklinik").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("Poliklinik").Range("C46").Value = tamsayi_kismi & "," & real_kismi
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    
    
        Ex.Sheets("Poliklinik").Range("R50").Value = tablo!fatura
    Ex.Sheets("Poliklinik").Range("R51").Value = tablo!sevk
    Ex.Sheets("Poliklinik").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Poliklinik").Range("C46").Value = "=P10"
    Ex.Sheets("Poliklinik").Range("L46").Value = "=P10-P15"
        'Ex.Sheets("Poliklinik").PrintOut 'Yazdýr
    ElseIf yazdirilacak = "Yolluk" Then
    Ex.Sheets("Yolluk").Range("T3").Value = tablo!adi
    Ex.Sheets("Yolluk").Range("T4").Value = tablo!tcno
    Ex.Sheets("Yolluk").Range("T5").Value = tablo!banka
    Ex.Sheets("Yolluk").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Yolluk").Range("T7").Value = tablo!vd
    Ex.Sheets("Yolluk").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Yolluk").Range("C46").Value - Ex.Sheets("Yolluk").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Yolluk").Range("C46").Value - Ex.Sheets("Yolluk").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Yolluk").Range("C46").Value - Ex.Sheets("Yolluk").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Yolluk").Range("C46").Value - Ex.Sheets("Yolluk").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Yolluk").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Yolluk").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Yolluk").Range("L46").Value = Round(Ex.Sheets("Yolluk").Range("C46").Value - Ex.Sheets("Yolluk").Range("I46").Value, 2)
    
    
    
    bul = InStr(Ex.Sheets("Yolluk").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Yolluk").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Yolluk").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Yolluk").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Yolluk").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Yolluk").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("Yolluk").Range("C46").Value = tamsayi_kismi & "," & real_kismi
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    Ex.Sheets("Yolluk").Range("R50").Value = tablo!gec_g_yol
    Ex.Sheets("Yolluk").Range("R51").Value = tablo!rayic
    Ex.Sheets("Yolluk").Range("R52").Value = tablo!sevk
    Ex.Sheets("Yolluk").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Yolluk").Range("C46").Value = "=P10"
    Ex.Sheets("Yolluk").Range("L46").Value = "=P10-P15"
    'Ex.Sheets("Yolluk").PrintOut 'Yazdýr
    
    
    ElseIf yazdirilacak = "YollukSurekli" Then
    Ex.Sheets("YollukSurekli").Range("T3").Value = tablo!adi
    Ex.Sheets("YollukSurekli").Range("T4").Value = tablo!tcno
    Ex.Sheets("YollukSurekli").Range("T5").Value = tablo!banka
    Ex.Sheets("YollukSurekli").Range("T6").Value = tablo!hesapno
    Ex.Sheets("YollukSurekli").Range("T7").Value = tablo!vd
    Ex.Sheets("YollukSurekli").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("YollukSurekli").Range("C46").Value - Ex.Sheets("YollukSurekli").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("YollukSurekli").Range("C46").Value - Ex.Sheets("YollukSurekli").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("YollukSurekli").Range("C46").Value - Ex.Sheets("YollukSurekli").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("YollukSurekli").Range("C46").Value - Ex.Sheets("YollukSurekli").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("YollukSurekli").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("YollukSurekli").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("YollukSurekli").Range("L46").Value = Round(Ex.Sheets("YollukSurekli").Range("C46").Value - Ex.Sheets("YollukSurekli").Range("I46").Value, 2)
    
    
    
    bul = InStr(Ex.Sheets("YollukSurekli").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("YollukSurekli").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("YollukSurekli").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("YollukSurekli").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("YollukSurekli").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("YollukSurekli").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("YollukSurekli").Range("C46").Value = tamsayi_kismi & "," & real_kismi
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    Ex.Sheets("YollukSurekli").Range("R50").Value = tablo!gec_g_yol
    Ex.Sheets("YollukSurekli").Range("R51").Value = tablo!rayic
    Ex.Sheets("YollukSurekli").Range("R52").Value = tablo!sevk
    Ex.Sheets("YollukSurekli").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("YollukSurekli").Range("C46").Value = "=P10"
    Ex.Sheets("YollukSurekli").Range("L46").Value = "=P10-P15"
    'Ex.Sheets("YollukSurekli").PrintOut 'Yazdýr
ElseIf yazdirilacak = "Yolluktedavi" Then
    Ex.Sheets("Yolluktedavi").Range("T3").Value = tablo!adi
    Ex.Sheets("Yolluktedavi").Range("T4").Value = tablo!tcno
    Ex.Sheets("Yolluktedavi").Range("T5").Value = tablo!banka
    Ex.Sheets("Yolluktedavi").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Yolluktedavi").Range("T7").Value = tablo!vd
    Ex.Sheets("Yolluktedavi").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Yolluktedavi").Range("C46").Value - Ex.Sheets("Yolluktedavi").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Yolluktedavi").Range("C46").Value - Ex.Sheets("Yolluktedavi").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Yolluktedavi").Range("C46").Value - Ex.Sheets("Yolluktedavi").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Yolluktedavi").Range("C46").Value - Ex.Sheets("Yolluktedavi").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Yolluktedavi").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Yolluktedavi").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Yolluktedavi").Range("L46").Value = Round(Ex.Sheets("Yolluktedavi").Range("C46").Value - Ex.Sheets("Yolluktedavi").Range("I46").Value, 2)
    
    
    
    bul = InStr(Ex.Sheets("Yolluktedavi").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Yolluktedavi").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Yolluktedavi").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Yolluktedavi").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Yolluktedavi").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Yolluktedavi").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("Yolluktedavi").Range("C46").Value = tamsayi_kismi & "," & real_kismi
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    Ex.Sheets("Yolluktedavi").Range("R50").Value = tablo!gec_g_yol
    Ex.Sheets("Yolluktedavi").Range("R51").Value = tablo!rayic
    Ex.Sheets("Yolluktedavi").Range("R52").Value = tablo!sevk
    Ex.Sheets("Yolluktedavi").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Yolluktedavi").Range("C46").Value = "=P10"
    Ex.Sheets("Yolluktedavi").Range("L46").Value = "=P10-P15"
    'Ex.Sheets("Yolluktedavi").PrintOut 'Yazdýr
    
    
ElseIf yazdirilacak = "huzurhakki" Then

    Ex.Sheets("Huzurhakki").Range("T3").Value = tablo!adi
    Ex.Sheets("Huzurhakki").Range("T4").Value = tablo!tcno
    Ex.Sheets("Huzurhakki").Range("T5").Value = tablo!banka
    Ex.Sheets("Huzurhakki").Range("T6").Value = tablo!hesapno
    Ex.Sheets("Huzurhakki").Range("T7").Value = tablo!vd
    Ex.Sheets("Huzurhakki").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("Huzurhakki").Range("C46").Value - Ex.Sheets("Huzurhakki").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("Huzurhakki").Range("C46").Value - Ex.Sheets("Huzurhakki").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("Huzurhakki").Range("C46").Value - Ex.Sheets("Huzurhakki").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("Huzurhakki").Range("C46").Value - Ex.Sheets("Huzurhakki").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Huzurhakki").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Huzurhakki").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("Huzurhakki").Range("L46").Value = Round(Ex.Sheets("Huzurhakki").Range("C46").Value - Ex.Sheets("Huzurhakki").Range("I46").Value, 2)
    
           
        
    
    bul = InStr(Ex.Sheets("Huzurhakki").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("Huzurhakki").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("Huzurhakki").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("Huzurhakki").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("Huzurhakki").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("Huzurhakki").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("Huzurhakki").Range("C46").Value = tamsayi_kismi & "," & real_kismi
    
    
    
    
    
        
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    
    Ex.Sheets("Huzurhakki").Range("R50").Value = tablo!bankalist
    Ex.Sheets("Huzurhakki").Range("R51").Value = tablo!cizelge
    Ex.Sheets("Huzurhakki").Range("R52").Value = tablo!bordro
    Ex.Sheets("Huzurhakki").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("Huzurhakki").Range("C46").Value = "=P10"
    Ex.Sheets("Huzurhakki").Range("L46").Value = "=P10-P15"




ElseIf yazdirilacak = "ogrprim" Then

    Ex.Sheets("ogrprim").Range("T3").Value = tablo!adi
    Ex.Sheets("ogrprim").Range("T4").Value = tablo!tcno
    Ex.Sheets("ogrprim").Range("T5").Value = tablo!banka
    Ex.Sheets("ogrprim").Range("T6").Value = tablo!hesapno
    Ex.Sheets("ogrprim").Range("T7").Value = tablo!vd
    Ex.Sheets("ogrprim").Range("P10").Value = tablo!borc
    
    '****************Bu satýrý  eklemeyi unutma*********************************************
    bul = InStr(Round(Ex.Sheets("ogrprim").Range("C46").Value - Ex.Sheets("ogrprim").Range("I46").Value, 2), ",")
    If bul = 0 Then
    tamsayi_kismi = Round(Ex.Sheets("ogrprim").Range("C46").Value - Ex.Sheets("ogrprim").Range("I46").Value, 2)
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Round(Ex.Sheets("ogrprim").Range("C46").Value - Ex.Sheets("ogrprim").Range("I46").Value, 2), 1, bul - 1)
    real_kismi = (Mid(Round(Ex.Sheets("ogrprim").Range("C46").Value - Ex.Sheets("ogrprim").Range("I46").Value, 2), bul + 1, 2))
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
    
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("ogrprim").Range("B62").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("ogrprim").Range("N62").Value = okunacakyeni_real
    
    Ex.Sheets("ogrprim").Range("L46").Value = Round(Ex.Sheets("ogrprim").Range("C46").Value - Ex.Sheets("ogrprim").Range("I46").Value, 2)
    
           
        
    
    bul = InStr(Ex.Sheets("ogrprim").Range("C46").Value, ",")
    If bul = 0 Then
    tamsayi_kismi = Ex.Sheets("ogrprim").Range("C46").Value
    real_kismi = 0
    Else
    tamsayi_kismi = Mid(Ex.Sheets("ogrprim").Range("C46").Value, 1, bul - 1)
    real_kismi = Mid(Ex.Sheets("ogrprim").Range("C46").Value, bul + 1, 3)
    If Len(real_kismi) <= 1 Then
    real_kismi = real_kismi & 0
    End If
    End If
        
    Call Virgulden_Once(tamsayi_kismi)
    Ex.Sheets("ogrprim").Range("B44").Value = okunacakyeni
    Call Virgulden_Sonra(real_kismi)
    Ex.Sheets("ogrprim").Range("L44").Value = okunacakyeni_real
    
    Ex.Sheets("ogrprim").Range("C46").Value = tamsayi_kismi & "," & real_kismi
    
    
    
    
    
        
'****************Bu satýrý  eklemeyi unutma*********************************************
    
    
    Ex.Sheets("ogrprim").Range("R50").Value = tablo!isimlist
    Ex.Sheets("ogrprim").Range("R51").Value = tablo!tahakkuk
    Ex.Sheets("ogrprim").PrintPreview
    'Formülü eski haline getirdim
    Ex.Sheets("ogrprim").Range("C46").Value = "=P10"
    Ex.Sheets("ogrprim").Range("L46").Value = "=P10-P15"

    
    
    End If
    
   
      
   'Ex.Sheets.PrintOut 'Yazdýr
   'Ex.Sheets("Eczane").PrintPreview 'Yazýcý önizlemesi için kullanýlabilir.
   Ex.ActiveWorkbook.Close True  'False deðeri yaptýklarýmýzýn kaydedilmemesi için
   Ex.Quit
Set Ex = Nothing

End Sub


Private Sub Option4_Click()
Call tablo_ac("Select * from Yolluk")
Call listele
Call yetki
yazdirilacak = "Yolluk"
End Sub



Private Sub Option9_Click()
Call tablo_ac("Select * from YollukSurekli")
Call listele
Call yetki
yazdirilacak = "YollukSurekli"
End Sub

Private Sub Timer1_Timer()
Me.Caption = "Nakit Kayýt ANA MENÜ      Bugün " & Now
End Sub

Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 240 Then
    trmLoad.Enabled = False
        Else
    t = t + 4
  End If
End Sub

Private Sub trmUnload_Timer()
SetTransparent hwnd, t
  If trmLoad.Enabled = True Then trmLoad.Enabled = False
  If t <= 0 Then
    trmUnload.Enabled = False
    Unload Me
        Else
    t = t - 4
  End If
  End Sub

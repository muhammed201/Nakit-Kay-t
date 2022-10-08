VERSION 5.00
Begin VB.Form eczane_duzelt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eczane Nakit Kaydý Düzeltme Ýþlemleri"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   Icon            =   "Eczane_duzelt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "ECZANE NAKÝT KAYDI DÜZELTME"
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
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   9375
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1800
         TabIndex        =   19
         Text            =   "1"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   6240
         TabIndex        =   18
         Text            =   "1"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   6120
         TabIndex        =   15
         Top             =   2760
         Width           =   2415
         Begin VB.CommandButton Command1 
            Caption         =   "Düzelt"
            Height          =   375
            Left            =   1200
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ýptal"
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   6240
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   6240
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   6240
         TabIndex        =   5
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fatura Say"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sevk Kaðýdý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   7
         Left            =   4680
         TabIndex        =   20
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kayýt Seçiniz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Borç"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   13
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vergi Dairesi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hesap Num."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   3
         Left            =   4800
         TabIndex        =   11
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banka Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   10
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.C. Num."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Eczane Adý"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1185
      End
   End
End
Attribute VB_Name = "eczane_duzelt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gtcno, gadi As String
Private Sub Combo1_Click()
Dim bul As Byte
bul = InStr(Combo1.Text, "-") - 2
Call tablo_ac("Select * from Eczane where tcno='" & Trim(Mid(Combo1.Text, 1, bul)) & "' and adi='" & Trim(Mid(Combo1.Text, bul + 3, Len(Combo1.Text) - bul + 3)) & "'")
gadi = tablo!adi
gtcno = tablo!tcno
Text1.Text = tablo!adi
Text2.Text = tablo!tcno
Text3.Text = tablo!banka
Text4.Text = tablo!hesapno
Text5.Text = tablo!vd
Text6.Text = tablo!borc & " YTL"
Text7.Text = tablo!fatura
Text8.Text = tablo!sevk
tablo.Close
End Sub

Private Sub Command1_Click()
On Local Error GoTo hata
Call tablo_ac("Select * from eczane where tcno='" & gtcno & "' and adi='" & gadi & "'")
tablo.Edit
tablo!adi = Text1.Text
tablo!tcno = Text2.Text
tablo!banka = Text3.Text
tablo!hesapno = Text4.Text
tablo!vd = Text5.Text
Text6.Text = Replace(Text6.Text, ".", ",")
If Right(Text6.Text, 3) = "YTL" Then
tablo!borc = Trim(Mid(Text6.Text, 1, Len(Text6.Text) - 3))
Else
tablo!borc = Trim(Text6.Text)
End If
tablo!fatura = Val(Text7.Text)
tablo!sevk = Val(Text8.Text)
tablo.Update
MsgBox "Günçelleþtirme baþarýlý bir þekilde gerçekleþtirilmiþtir", vbInformation
Call yenile
hata:
If Err = "3021" Then
MsgBox "Düzeltilecek kayýt bulunamadý", vbCritical
End If
End Sub

Private Sub Command2_Click()
Unload Me
AnaMenu.Show
End Sub

Private Sub Form_Load()
Call veri_ac(False, False)
Call tablo_ac("Select * from eczane")
Combo1.Clear
Do While Not tablo.EOF
Combo1.AddItem tablo!tcno & " - " & tablo!adi
tablo.MoveNext
Loop
tablo.Close
End Sub

Sub yenile()
Call veri_ac(False, False)
Call tablo_ac("Select * from Eczane")
Combo1.Clear
Do While Not tablo.EOF
Combo1.AddItem tablo!tcno & " - " & tablo!adi
tablo.MoveNext
Loop
tablo.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnaMenu.Show
End Sub


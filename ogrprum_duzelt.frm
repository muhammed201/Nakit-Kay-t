VERSION 5.00
Begin VB.Form ogrprim_duzelt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Öðrenci Sigorta Primleri  Kaydý Düzeltme"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   Icon            =   "ogrprum_duzelt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Öðrenci Sigorta Primleri Düzeltme"
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
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ýptal"
         Height          =   615
         Left            =   6480
         TabIndex        =   10
         Top             =   2520
         Width           =   975
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
         TabIndex        =   9
         Top             =   1680
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
         Left            =   6600
         TabIndex        =   8
         Text            =   "1"
         Top             =   1680
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
         Left            =   6600
         TabIndex        =   7
         Text            =   "1"
         Top             =   1080
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
         TabIndex        =   6
         Top             =   1080
         Width           =   2175
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
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Düzelt"
         Height          =   615
         Left            =   7440
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
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
         Left            =   6600
         TabIndex        =   3
         Top             =   480
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
         Left            =   1800
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text9 
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
         TabIndex        =   1
         Top             =   2880
         Width           =   2175
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
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adý-Soyadý"
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
         TabIndex        =   18
         Top             =   480
         Width           =   1170
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
         TabIndex        =   17
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahakkuk"
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
         Left            =   4920
         TabIndex        =   16
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ýsim Listesi"
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
         Left            =   4920
         TabIndex        =   15
         Top             =   1680
         Width           =   1155
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1290
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
         Index           =   6
         Left            =   4920
         TabIndex        =   13
         Top             =   480
         Width           =   495
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
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1080
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
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1380
      End
   End
End
Attribute VB_Name = "ogrprim_duzelt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gtcno, gadi As String
Private Sub Combo1_Click()
Dim bul As Byte
bul = InStr(Combo1.Text, "-") - 2
Call tablo_ac("Select * from ogrprim where tcno='" & Trim(Mid(Combo1.Text, 1, bul)) & "' and adi='" & Trim(Mid(Combo1.Text, bul + 3, Len(Combo1.Text) - bul + 3)) & "'")
gadi = tablo!adi
gtcno = tablo!tcno
Text1.Text = tablo!adi
Text2.Text = tablo!tcno
Text3.Text = tablo!tahakkuk
Text4.Text = tablo!isimlist
Text5.Text = tablo!hesapno

Text7.Text = tablo!borc & " YTL"
Text8.Text = tablo!banka
Text9.Text = tablo!vd
tablo.Close
End Sub

Private Sub Command1_Click()
On Local Error GoTo hata:
Call tablo_ac("Select * from ogrprim where tcno='" & gtcno & "' and adi='" & gadi & "'")
tablo.Edit
tablo!adi = Text1.Text
tablo!tcno = Text2.Text
tablo!banka = Text8.Text
tablo!hesapno = Text5.Text
tablo!vd = Text9.Text
Text7.Text = Replace(Text7.Text, ".", ",")
If Right(Text7.Text, 3) = "YTL" Then
tablo!borc = Trim(Mid(Text7.Text, 1, Len(Text7.Text) - 3))
Else
tablo!borc = Trim(Text7.Text)
End If
tablo!tahakkuk = Text3.Text
tablo!isimlist = Text4.Text

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
Call tablo_ac("Select * from ogrprim")
Combo1.Clear
Do While Not tablo.EOF
Combo1.AddItem tablo!tcno & " - " & tablo!adi
tablo.MoveNext
Loop
tablo.Close
End Sub

Sub yenile()
Call veri_ac(False, False)
Call tablo_ac("Select * from ogrprim")
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


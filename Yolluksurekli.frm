VERSION 5.00
Begin VB.Form Yolluksurekli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yolluk(Sürekli) Nakit Kaydý"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "Yolluksurekli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "YOLLUK(SÜREKLÝ) NAKÝT KAYDI"
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
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9255
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
         TabIndex        =   4
         Top             =   2880
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
         TabIndex        =   3
         Top             =   2280
         Width           =   2175
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
         TabIndex        =   5
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Kaydet"
         Height          =   615
         Left            =   7320
         TabIndex        =   9
         Top             =   2760
         Width           =   1455
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
         Left            =   6600
         TabIndex        =   6
         Text            =   "1"
         Top             =   1080
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
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
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
         Left            =   6600
         TabIndex        =   8
         Text            =   "1"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ýptal"
         Height          =   615
         Left            =   6480
         TabIndex        =   10
         Top             =   2760
         Width           =   855
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
         TabIndex        =   20
         Top             =   2880
         Width           =   1380
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
         TabIndex        =   19
         Top             =   2280
         Width           =   1080
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
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rayiç"
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
         Left            =   4920
         TabIndex        =   17
         Top             =   2280
         Width           =   615
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
         TabIndex        =   16
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Geç.G.Yol."
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
         Width           =   1125
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
         Index           =   2
         Left            =   4920
         TabIndex        =   14
         Top             =   1080
         Width           =   1260
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
         TabIndex        =   13
         Top             =   1080
         Width           =   1020
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
         TabIndex        =   12
         Top             =   480
         Width           =   1170
      End
   End
End
Attribute VB_Name = "Yolluksurekli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call kayit
End Sub

Private Sub Command2_Click()
Unload Me
AnaMenu.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnaMenu.Show
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &H80000018
End Sub


Private Sub Text2_GotFocus()
Text2.BackColor = &H80000018
End Sub

Private Sub Text3_GotFocus()
Text3.BackColor = &H80000018
End Sub

Private Sub Text4_GotFocus()
Text4.BackColor = &H80000018
End Sub

Private Sub Text5_GotFocus()
Text5.BackColor = &H80000018
End Sub

Private Sub Text6_GotFocus()
Text6.BackColor = &H80000018
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &H80000014
End Sub

Private Sub Text2_LostFocus()
Text2.BackColor = &H80000014
End Sub

Private Sub Text3_LostFocus()
Text3.BackColor = &H80000014
End Sub

Private Sub Text4_LostFocus()
Text4.BackColor = &H80000014
End Sub

Private Sub Text5_LostFocus()
Text5.BackColor = &H80000014
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = &H80000014
End Sub

Sub kayit()
Call veri_ac(False, False)
Call tablo_ac("Select * from Yolluksurekli")
tablo.AddNew

If Text1.Text = Empty Then
MsgBox "Boþ alanlar var", vbCritical
Exit Sub
Else
tablo!adi = Text1.Text
End If

If Text2.Text = Empty Then
MsgBox "Boþ alanlar var", vbCritical
Exit Sub
Else
tablo!tcno = Text2.Text
End If

tablo!sevk = Text3.Text
tablo!gec_g_yol = Text4.Text
tablo!hesapno = Text5.Text
tablo!rayic = Text6.Text

Text7.Text = Replace(Text7.Text, ".", ",")

If Text7.Text = Empty Then
Text7.Text = 0
End If

tablo!borc = Text7.Text
tablo!banka = Text8.Text
tablo!vd = Text9.Text

tablo.Update
MsgBox "Kayýt baþarýyla gerçekleþmiþtir", vbInformation
Call temizle
End Sub

Sub temizle()
Text1.Text = Empty
Text2.Text = Empty
Text5.Text = Empty
Text7.Text = Empty
Text8.Text = Empty
End Sub

Private Sub Text7_GotFocus()
Text5.BackColor = &H80000018
End Sub

Private Sub Text7_LostFocus()
Text1.BackColor = &H80000014
End Sub

Private Sub Text8_GotFocus()
Text5.BackColor = &H80000018
End Sub

Private Sub Text8_LostFocus()
Text1.BackColor = &H80000014
End Sub

Private Sub Text9_GotFocus()
Text5.BackColor = &H80000018
End Sub

Private Sub Text9_LostFocus()
Text1.BackColor = &H80000014
End Sub

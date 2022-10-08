VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form tumliste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tüm Nakitlerin Listesi"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11310
   Icon            =   "tumliste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Seçenekler"
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
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   1560
         Picture         =   "tumliste.frx":33E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   600
         Picture         =   "tumliste.frx":166E6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   4575
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483644
      ForeColor       =   -2147483647
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex2 
      Height          =   4575
      Left            =   7680
      TabIndex        =   4
      Top             =   1320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483644
      ForeColor       =   -2147483647
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex3 
      Height          =   4575
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483644
      ForeColor       =   -2147483647
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Flex4 
      Height          =   4575
      Left            =   7680
      TabIndex        =   6
      Top             =   6000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColor       =   -2147483644
      ForeColor       =   -2147483647
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "tumliste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bir()
Dim i As Integer
On Local Error Resume Next
Flex.Rows = 1
Flex.Cols = 8
Flex.TextMatrix(0, 0) = "ECZANE ADI"
Flex.ColWidth(0) = 150 * 20
Flex.TextMatrix(0, 1) = "BANKA"
Flex.ColWidth(1) = 150 * 22
Flex.TextMatrix(0, 2) = "T.C.NO"
Flex.ColWidth(2) = 150 * 12
Flex.TextMatrix(0, 3) = "HESAP NO"
Flex.ColWidth(3) = 150 * 10
Flex.TextMatrix(0, 4) = "VERGÝ DAÝRESÝ"
Flex.ColWidth(4) = 150 * 15
Flex.TextMatrix(0, 5) = "BORÇ"
Flex.ColWidth(5) = 150 * 10
Flex.TextMatrix(0, 6) = "FATURA"
Flex.ColWidth(6) = 150 * 7
Flex.TextMatrix(0, 7) = "SEVK"
Flex.ColWidth(7) = 150 * 10



Call veri_ac(False, False)
Call tablo_ac("select * from eczane")
i = 0
Do While Not tablo.EOF
i = i + 1
Flex.AddItem ""
Flex.TextMatrix(i, 0) = tablo!adi
Flex.TextMatrix(i, 1) = tablo!banka
Flex.TextMatrix(i, 2) = tablo!tcno
Flex.TextMatrix(i, 3) = tablo!hesapno
Flex.TextMatrix(i, 4) = tablo!vd
Flex.TextMatrix(i, 5) = tablo!borc & " YTL"
Flex.TextMatrix(i, 6) = tablo!fatura
Flex.TextMatrix(i, 7) = tablo!sevk
tablo.MoveNext
Loop
tablo.Close
'veri.Close
End Sub

Sub iki()
Dim I2 As Integer
On Local Error Resume Next
Flex2.Rows = 1
Flex2.Cols = 8
Flex2.TextMatrix(0, 0) = "HASTANE ADI"
Flex2.ColWidth(0) = 150 * 20
Flex2.TextMatrix(0, 1) = "BANKA"
Flex2.ColWidth(1) = 150 * 22
Flex2.TextMatrix(0, 2) = "T.C.NO"
Flex2.ColWidth(2) = 150 * 12
Flex2.TextMatrix(0, 3) = "HESAP NO"
Flex2.ColWidth(3) = 150 * 10
Flex2.TextMatrix(0, 4) = "VERGÝ DAÝRESÝ"
Flex2.ColWidth(4) = 150 * 15
Flex2.TextMatrix(0, 5) = "BORÇ"
Flex2.ColWidth(5) = 150 * 10
Flex2.TextMatrix(0, 6) = "FATURA"
Flex2.ColWidth(6) = 150 * 7
Flex2.TextMatrix(0, 7) = "SEVK"
Flex2.ColWidth(7) = 150 * 10

Call veri_ac(False, False)
Call tablo_ac("select * from hastane")
I2 = 0
Do While Not tablo.EOF
I2 = I2 + 1
Flex2.AddItem ""
Flex2.TextMatrix(I2, 0) = tablo!adi
Flex2.TextMatrix(I2, 1) = tablo!banka
Flex2.TextMatrix(I2, 2) = tablo!tcno
Flex2.TextMatrix(I2, 3) = tablo!hesapno
Flex2.TextMatrix(I2, 4) = tablo!vd
Flex2.TextMatrix(I2, 5) = tablo!borc & " YTL"
Flex2.TextMatrix(I2, 6) = tablo!fatura
Flex2.TextMatrix(I2, 7) = tablo!sevk
tablo.MoveNext
Loop
tablo.Close
'veri.Close
End Sub

Sub uc()
Dim I3 As Integer
On Local Error Resume Next
Flex3.Rows = 1
Flex3.Cols = 8
Flex3.TextMatrix(0, 0) = "POLÝKLÝNÝK ADI"
Flex3.ColWidth(0) = 150 * 20
Flex3.TextMatrix(0, 1) = "BANKA"
Flex3.ColWidth(1) = 150 * 22
Flex3.TextMatrix(0, 2) = "T.C.NO"
Flex3.ColWidth(2) = 150 * 12
Flex3.TextMatrix(0, 3) = "HESAP NO"
Flex3.ColWidth(3) = 150 * 10
Flex3.TextMatrix(0, 4) = "VERGÝ DAÝRESÝ"
Flex3.ColWidth(4) = 150 * 15
Flex3.TextMatrix(0, 5) = "BORÇ"
Flex3.ColWidth(5) = 150 * 10
Flex3.TextMatrix(0, 6) = "FATURA"
Flex3.ColWidth(6) = 150 * 7
Flex3.TextMatrix(0, 7) = "SEVK"
Flex3.ColWidth(7) = 150 * 10

Call veri_ac(False, False)
Call tablo_ac("select * from poliklinik")
I3 = 0
Do While Not tablo.EOF
I3 = I3 + 1
Flex3.AddItem ""
Flex3.TextMatrix(I3, 0) = tablo!adi
Flex3.TextMatrix(I3, 1) = tablo!banka
Flex3.TextMatrix(I3, 2) = tablo!tcno
Flex3.TextMatrix(I3, 3) = tablo!hesapno
Flex3.TextMatrix(I3, 4) = tablo!vd
Flex3.TextMatrix(I3, 5) = tablo!borc & " YTL"
Flex3.TextMatrix(I3, 6) = tablo!fatura
Flex3.TextMatrix(I3, 7) = tablo!sevk
tablo.MoveNext
Loop
tablo.Close
'veri.Close
End Sub
Sub dort()
Dim i As Integer
On Local Error Resume Next
Flex4.Rows = 1
Flex4.Cols = 9
Flex4.TextMatrix(0, 0) = "ADI-SOYADI"
Flex4.ColWidth(0) = 150 * 20
Flex4.TextMatrix(0, 1) = "T.C. NO"
Flex4.ColWidth(1) = 150 * 15
Flex4.TextMatrix(0, 2) = "HESAP NO"
Flex4.ColWidth(2) = 150 * 12
Flex4.TextMatrix(0, 3) = "BANKA"
Flex4.ColWidth(3) = 150 * 10
Flex4.TextMatrix(0, 4) = "VERGÝ DAÝRESÝ"
Flex4.ColWidth(4) = 150 * 15
Flex4.TextMatrix(0, 5) = "BORÇ"
Flex4.ColWidth(5) = 150 * 10
Flex4.TextMatrix(0, 6) = "GEÇ.G.YOL"
Flex4.ColWidth(5) = 150 * 10
Flex4.TextMatrix(0, 7) = "RAYÝÇ"
Flex4.ColWidth(5) = 150 * 10
Flex4.TextMatrix(0, 8) = "SEVK KAÐ."
Flex4.ColWidth(5) = 150 * 10


Call veri_ac(False, False)
Call tablo_ac("select * from yolluk")
i = 0
Do While Not tablo.EOF
i = i + 1
Flex4.AddItem ""
Flex4.TextMatrix(i, 0) = tablo!adi
Flex4.TextMatrix(i, 1) = tablo!tcno
Flex4.TextMatrix(i, 2) = tablo!hesapno
Flex4.TextMatrix(i, 3) = tablo!banka
Flex4.TextMatrix(i, 4) = tablo!vd
Flex4.TextMatrix(i, 5) = tablo!borc & " YTL"
Flex4.TextMatrix(i, 6) = tablo!gec_g_yol
Flex4.TextMatrix(i, 7) = tablo!rayic
Flex4.TextMatrix(i, 8) = tablo!sevk
tablo.MoveNext
Loop
tablo.Close
End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Frame1.Left = Flex.Left
Frame1.Width = Screen.Width - Screen.Width / 10
Call bir
Call iki
Call uc
Call dort
End Sub


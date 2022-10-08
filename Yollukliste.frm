VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Yollukliste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yolluk(Geçici) Nakit Listesi"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11505
   Icon            =   "Yollukliste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   11505
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13215
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   1560
         Picture         =   "Yollukliste.frx":33E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   600
         Picture         =   "Yollukliste.frx":166E6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   8895
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15690
      _Version        =   393216
      Cols            =   9
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
Attribute VB_Name = "Yollukliste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim a, i As Long
If Flex.Rows = 1 Then ' Eðer msflexgrid boþsa aktarma yapma
Title = "Hata !"
Msg = "Microsoft Excel' e aktarýlacak herhangi bir kayýt bulunamadý."
Answer = MsgBox(Msg, vbCritical, Title)
Exit Sub
End If
Screen.MousePointer = 11
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets(1)


For a = 1 To Flex.Cols
For i = 1 To Flex.Rows
xlSheet.Cells(1, a).Font.Bold = True ' Excel' e aktarýlan kayýtlarýn ilk sütunu baþlýklar olduðu için karakterleri bold
xlSheet.Cells(i, a) = Flex.TextMatrix(i - 1, a - 1)
'xlSheet.Cells(i, 3).NumberFormat = "#,##0.00" 'bexcel tarafna aktarýlan kaydýn formatý
Next i
Next a

Screen.MousePointer = 0
xlBook.Application.Visible = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
Flex.Width = Screen.Width - 20
Flex.Left = 20
On Local Error Resume Next
Flex.Rows = 1
Flex.Cols = 9
Flex.TextMatrix(0, 0) = "ADI-SOYADI"
Flex.ColWidth(0) = 150 * 20
Flex.TextMatrix(0, 1) = "T.C. NO"
Flex.ColWidth(1) = 150 * 15
Flex.TextMatrix(0, 2) = "HESAP NO"
Flex.ColWidth(2) = 150 * 12
Flex.TextMatrix(0, 3) = "BANKA"
Flex.ColWidth(3) = 150 * 10
Flex.TextMatrix(0, 4) = "VERGÝ DAÝRESÝ"
Flex.ColWidth(4) = 150 * 15
Flex.TextMatrix(0, 5) = "BORÇ"
Flex.ColWidth(5) = 150 * 10
Flex.TextMatrix(0, 6) = "GEÇ.G.YOL"
Flex.ColWidth(5) = 150 * 10
Flex.TextMatrix(0, 7) = "RAYÝÇ"
Flex.ColWidth(5) = 150 * 10
Flex.TextMatrix(0, 8) = "SEVK KAÐ."
Flex.ColWidth(5) = 150 * 10


Call veri_ac(False, False)
Call tablo_ac("select * from yolluk")
i = 0
Do While Not tablo.EOF
i = i + 1
Flex.AddItem ""
Flex.TextMatrix(i, 0) = tablo!adi
Flex.TextMatrix(i, 1) = tablo!tcno
Flex.TextMatrix(i, 2) = tablo!hesapno
Flex.TextMatrix(i, 3) = tablo!banka
Flex.TextMatrix(i, 4) = tablo!vd
Flex.TextMatrix(i, 5) = tablo!borc & " YTL"
Flex.TextMatrix(i, 6) = tablo!gec_g_yol
Flex.TextMatrix(i, 7) = tablo!rayic
Flex.TextMatrix(i, 8) = tablo!sevk
tablo.MoveNext
Loop
tablo.Close
'veri.Close
Frame1.Left = Flex.Left
Frame1.Width = Flex.Width
End Sub





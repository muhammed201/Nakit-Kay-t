VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form Hastaneliste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hastane Nakit Listesi"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   Icon            =   "Hastaneliste.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10530
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   600
         Picture         =   "Hastaneliste.frx":33E2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Height          =   855
         Left            =   1560
         Picture         =   "Hastaneliste.frx":3CAC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   8895
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   15690
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
Attribute VB_Name = "Hastaneliste"
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
'xlSheet.Cells(i, 3).NumberFormat = "#,##0.00" ' excel tarafna aktarýlan kaydýn formatý
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
Flex.Left = 100
Flex.Width = Screen.Width - 100
On Local Error Resume Next
Flex.Rows = 1
Flex.Cols = 8
Flex.TextMatrix(0, 0) = "HASTANE ADI"
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
Flex.ColWidth(7) = 150 * 5


Call veri_ac(False, False)
Call tablo_ac("select * from hastane")
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
Frame1.Left = Flex.Left
Frame1.Width = Flex.Width
End Sub



Attribute VB_Name = "Veritabani"
Global veri As Database
Global tablo As Recordset
Global satir, sutun As String
Global t As Byte
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


Sub veri_ac(X1 As Boolean, X2 As Boolean)
Dim sifre
sifre = ";pwd=" & Chr(49) & Chr(52) & Chr(53) & Chr(51)
Set veri = Workspaces(0).OpenDatabase(App.Path & "\veritabani.mdb", X1, X2, sifre)
End Sub


Sub tablo_ac(sql As String)
Set tablo = veri.OpenRecordset(sql)
End Sub


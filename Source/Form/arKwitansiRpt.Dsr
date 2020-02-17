VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arKwitansiRpt 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arKwitansiRpt.dsx":0000
   PaletteMode     =   2  'Custom
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arKwitansiRpt.dsx":628A
End
Attribute VB_Name = "arKwitansiRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim terbilang As New CRUFLFungsiku.Konversi

Private Sub ActiveReport_ReportStart()
lblCompany1.Caption = arKwitansiRpt.adoKu.Recordset.Fields("CmnyName").value
lblalamat.Caption = arKwitansiRpt.adoKu.Recordset.Fields("Address1").value
lblKota = arKwitansiRpt.adoKu.Recordset.Fields("City").value
lblFax = arKwitansiRpt.adoKu.Recordset.Fields("Faximale").value & " " & arKwitansiRpt.adoKu.Recordset.Fields("Phone1").value

If MenuFrm.simage_name = "" Then
    foto = ""
Else
    foto = App.Path & "\images\" & MenuFrm.simage_name
    Image1.Picture = LoadPicture(foto)
End If

If MenuFrm.sis_image = 0 Then
    Image1.Visible = False
    lblCompany1.Left = 270
    lblalamat.Left = 270
    lblKota.Left = 270
    lblFax.Left = 270
    Label37.Left = 270
End If

End Sub

Private Sub Detail_Format()
With arKwitansiRpt.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
    'sItmCount = sItmCount + .Fields("jumlah").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_BeforePrint()
lblterbilang.Caption = terbilang.terbilang(Field10)

End Sub

Private Sub GroupHeader1_BeforePrint()
Field12.text = arKwitansiRpt.adoKu.Recordset.Fields("kodecustomer").value + "-" + arKwitansiRpt.adoKu.Recordset.Fields("namacustomer").value
sno = 0
End Sub


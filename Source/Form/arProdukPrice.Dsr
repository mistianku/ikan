VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arProdukPrice 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arProdukPrice.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arProdukPrice.dsx":628A
End
Attribute VB_Name = "arProdukPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim skodeproduk As String

Private Sub ActiveReport_ReportStart()
lblCompany1.Font.Size = 10
lblCompany2.Font.Size = 7
lblCompany2.Font.Size = 7

If MenuFrm.simage_name = "" Then
    foto = ""
Else
    foto = App.Path & "\images\" & MenuFrm.simage_name
    Image1.Picture = LoadPicture(foto)
End If

If MenuFrm.sis_image = 0 Then
    Image1.Visible = False
    lblCompany1.Left = 90
    lblCompany2.Left = 90
    lblCompany3.Left = 90
End If

End Sub

Private Sub Detail_Format()
With arProdukPrice.adoKu.Recordset

If Not .EOF Then
    If skodeproduk <> .Fields("kodeproduk").value Then
        sno = sno + 1
        skodeproduk = .Fields("kodeproduk").value
        txtNo.Visible = True
        txtkodeproduk.Visible = True
        txtnamaproduk.Visible = True
    Else
        txtNo.Visible = False
        txtkodeproduk.Visible = False
        txtnamaproduk.Visible = False
    End If
   txtNo.text = sno
End If
End With
End Sub

Private Sub GroupFooter1_Format()
'txtpotonganfaktur.Text = sItmCount
'txttotdebet.Text = formatRupiah(samtdebet)
'txttotkredit.Text = formatRupiah(samtkredit)

End Sub

Private Sub GroupHeader2_Format()
sno = 0
With arProdukPrice.adoKu.Recordset

If Not .EOF Then
    lblbrand.Caption = "Brand   : " & .Fields("kodebrand").value & "-" & .Fields("namabrand").value

End If
End With
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = "Tanggal : " & Format(Now(), "DD-MM-YYYY HH:MM")
End Sub

'Private Sub PageHeader_BeforePrint()
'txtPeriode.Text = "Periode : " & arGLTrailBalance.adoKu.Recordset.Fields("yop").value & " - " & arGLTrailBalance.adoKu.Recordset.Fields("mop").value
'End Sub


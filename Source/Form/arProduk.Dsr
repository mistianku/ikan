VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arProduk 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arProduk.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arProduk.dsx":628A
End
Attribute VB_Name = "arProduk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim samtdebet As Double
Dim samtkredit As Double

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
With arProduk.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
    If .Fields("aktif").value = "1" Then
        txtaktif.text = "Y"
    Else
        txtaktif.text = "N"
    End If
    txtKategori = .Fields("kodekategori").value & "-" & .Fields("namakategori").value
    txtfungsi = .Fields("kodefungsi").value & "-" & .Fields("namafungsi").value
    txtumo1 = .Fields("uom1").value & "-" & .Fields("uom1sat").value
    txtumo2 = .Fields("uom2").value & "-" & .Fields("uom2sat").value
    txtumo3 = .Fields("umo3").value & "-" & .Fields("uom3sat").value
Else
End If
End With
End Sub


Private Sub GroupHeader1_Format()
With arProduk.adoKu.Recordset
    If Not .EOF Then
    lblbrand.Caption = "Brand      : " & .Fields("kodebrand").value & "-" & .Fields("namabrand").value
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


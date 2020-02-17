VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arMutasiStockFrm 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arMutasiStockFrm.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arMutasiStockFrm.dsx":628A
End
Attribute VB_Name = "arMutasiStockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sstockm, sstockk, sstockakhir As Double
Dim sItmCount As Double

Private Sub ActiveReport_ReportStart()
'lblCompany1.Font.Size = 10
'lblCompany2.Font.Size = 7
'lblCompany2.Font.Size = 7
lblCompany1.Caption = arMutasiStockFrm.adoKu.Recordset.Fields("CmnyName").value
lblalamat.Caption = arMutasiStockFrm.adoKu.Recordset.Fields("Address1").value
lblKota = arMutasiStockFrm.adoKu.Recordset.Fields("City").value
lblFax = arMutasiStockFrm.adoKu.Recordset.Fields("Faximale").value & " " & arMutasiStockFrm.adoKu.Recordset.Fields("Phone1").value


If MenuFrm.simage_name = "" Then
    foto = ""
Else
    foto = App.Path & "\images\" & MenuFrm.simage_name
    Image1.Picture = LoadPicture(foto)
    Image1.SizeMode = ddSMStretch
End If

If MenuFrm.sis_image = 0 Then
    Image1.Visible = False
    lblCompany1.Left = 90
    lblalamat.Left = 90
    lblKota.Left = 90
    lblFax.Left = 90
End If

End Sub

Private Sub Detail_BeforePrint()
With arMutasiStockFrm.adoKu.Recordset
'    If .Fields("txtdesc").value = "Saldo Awal" Then
'        sno = 0
'    End If
'    stockm.Text = Format(stockm.Text, "###,###.#0")
'    stockk.Text = Format(stockk.Text, "###,###.#0")
'    sisa.Text = Format(sisa.Text, "###,###.#0")
End With
End Sub

Private Sub Detail_Format()
With arMutasiStockFrm.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
'    sItmCount = sItmCount + .Fields("stockm").value
'    sstockm = sstockm + .Fields("stockm").value
'    sstockk = sstockk + .Fields("stockk").value
'    sstockakhir = sstockakhir + .Fields("cstock").value
'    sisa.Text = sstockakhir
'    If .Fields("txtdesc").value = "Saldo Awal" Then
'        txtNo.Visible = False
'        txtketeranganku.Visible = False
'        txttanggal.Visible = False
'        txtketerangan.Visible = False
'    Else
'        txtNo.Visible = True
'        txtketeranganku.Visible = True
'        txttanggal.Visible = True
'        txtketerangan.Visible = True
'    End If
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
'    Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
'    Field13.Text = .Fields("keterangan").value
'    Field14.Text = .Fields("referensi").value
End If
End With
End Sub

Private Sub GroupFooter1_BeforePrint()
'stkm.Text = Format(sstockm, "###,###,###.#0")
'stkk.Text = Format(sstockk, "###,###,###.#0")
'stockakh.Text = Format(sstockakhir, "###,###,###.#0")
End Sub

Private Sub GroupFooter1_Format()
'With arBeliFrm.adoKu.Recordset
'   Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
'   Field13.Text = .Fields("keterangan").value
'   Field14.Text = .Fields("referensi").value
'End With
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub


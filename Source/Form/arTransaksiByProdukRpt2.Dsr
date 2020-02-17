VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTransaksiByProdukRpt2 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arTransaksiByProdukRpt2.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arTransaksiByProdukRpt2.dsx":628A
End
Attribute VB_Name = "arTransaksiByProdukRpt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim sno2 As Integer
Dim snodok As String
Dim sItmCount As Double

Private Sub ActiveReport_ReportStart()
'lblCompany1.Font.Size = 10
'lblCompany2.Font.Size = 7
'lblCompany2.Font.Size = 7
lblCompany1.Caption = arTransaksiByProdukRpt2.adoKu.Recordset.Fields("CmnyName").value
lblalamat.Caption = arTransaksiByProdukRpt2.adoKu.Recordset.Fields("Address1").value
lblKota = arTransaksiByProdukRpt2.adoKu.Recordset.Fields("City").value
lblFax = arTransaksiByProdukRpt2.adoKu.Recordset.Fields("Faximale").value & " " & arTransaksiByProdukRpt2.adoKu.Recordset.Fields("Phone1").value

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

Private Sub Detail_Format()
With arTransaksiByProdukRpt2.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    sno2 = sno2 + 1
    txtNo.text = sno
'    sItmCount = sItmCount + .Fields("jumlah").value
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
'    Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
'   Field13.Text = .Fields("keterangan").value
'   Field14.Text = .Fields("referensi").value
End If
End With
End Sub

Private Sub GroupFooter1_Format()
'With arBeliFrm.adoKu.Recordset
'   Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
'   Field13.Text = .Fields("keterangan").value
'   Field14.Text = .Fields("referensi").value
'End With
End Sub

Private Sub GroupFooter2_AfterPrint()
'        If sno = 1 Then
'            arTransaksiByProdukRpt.GroupFooter2.Visible = False
'        End If
End Sub

Private Sub GroupFooter2_BeforePrint()
        lblfooter1 = " Sub Total : (" & lblheader2 & ")"
        lblfooter2 = " Sub Total : (" & lblheader & ")"
        sno = 0
End Sub

Private Sub GroupFooter2_Format()
        If sno = 1 Then
            sno = 0
            arTransaksiByProdukRpt.GroupFooter2.Visible = False
        End If
        
End Sub

Private Sub GroupHeader1_BeforePrint()
With arTransaksiByProdukRpt2.adoKu.Recordset
    
'   If .Fields("keysortby").value = 1 Then
        lblheader = "Tanggal  : " & Format(arTransaksiByProdukRpt2.adoKu.Recordset.Fields("tgldokumen").value, "dd-mm-yyyy")
        lblfooter1 = " Sub Total " & lblheader

'   Else
'        lblheader = "Supplier : " & arMasukLainRpt.adoKu.Recordset.Fields("namacustomer").value
'   End If
   
End With

sno = 0

End Sub

Private Sub GroupHeader2_Format()
With arTransaksiByProdukRpt2.adoKu.Recordset
    
'   If .Fields("keysortby").value = 1 Then
        lblheader = "Tanggal  : " & Format(arTransaksiByProdukRpt2.adoKu.Recordset.Fields("tgldokumen").value, "dd-mm-yyyy")
        lblheader2 = ToText(arTransaksiByProdukRpt2.adoKu.Recordset.Fields("sortdesc"))

'   Else
'        lblheader = "Supplier : " & arMasukLainRpt.adoKu.Recordset.Fields("namacustomer").value
'   End If
End With
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub

Private Sub ReportFooter_BeforePrint()
lblnoQ = "#" & sno2
End Sub


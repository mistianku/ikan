VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arlhppForm 
   Caption         =   "Kwitansi"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   Icon            =   "arlhppForm141209.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "arlhppForm141209.dsx":C84A
End
Attribute VB_Name = "arlhppForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double

Private Sub ActiveReport_ReportStart()
''lblCompany1.Font.Size = 10
''lblCompany2.Font.Size = 7
''lblCompany2.Font.Size = 7
'lblCompany1.Caption = arlhppForm.adoKu.Recordset.Fields("CmnyName").value
'lblalamat.Caption = arlhppForm.adoKu.Recordset.Fields("Address1").value
'lblKota = arlhppForm.adoKu.Recordset.Fields("City").value
'lblFax = arlhppForm.adoKu.Recordset.Fields("Faximale").value & " " & arlhppForm.adoKu.Recordset.Fields("Phone1").value
End Sub

Private Sub Detail_Format()
 
With arlhppForm.adoKu.Recordset

If Not .EOF Then

    If txtnolhpp1.Text = txtkolektor1.Text Then
        txtnolhpp1.Font.Bold = True
        txtnilailhpp1.Visible = False
        txtnolhpp1.Visible = False
    Else
        txtnolhpp1.Font.Bold = False
        txtnilailhpp1.Visible = True
        txtnolhpp1.Visible = True
    End If
    If Field1.Text = Field3.Text Then
        Field3.Visible = False
        Field2.Visible = False
    Else
        Field3.Visible = True
        Field2.Visible = True
    End If
    
    If txtkolektor1.Text = "-" Then
        txtkolektor1.Visible = False
        txtnilailhpp1.Visible = False
    Else
        txtkolektor1.Visible = True
        'txtnilailhpp1.Visible = True
    End If
    
    If Field1.Text = "-" Then
        
        Field1.Visible = False
        Field2.Visible = False
        
    Else
        
        Field1.Visible = True
        'Field2.Visible = True
    End If
    
'    sno = sno + 1
'    txtNo.Text = sno
  ''  sItmCount = sItmCount + .Fields("jumlah").value
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
    Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
   Field13.Text = .Fields("keterangan").value
   Field14.Text = .Fields("referensi").value
End If
End With
End Sub

Private Sub GroupFooter1_Format()
'With arlhppForm.adoKu.Recordset
'   Field12.Text = .Fields("kodegudang").value & " - " & .Fields("namagudang").value
'   Field13.Text = .Fields("keterangan").value
'   Field14.Text = .Fields("referensi").value
'End With
End Sub

Private Sub PageFooter_BeforePrint()
'lblHalaman.Caption = "Halaman ke : " & txtPageNumber.Text & " dari " & Me.Pages.Count
'lblUser.Caption = "User : " & MenuFrm.sUserID
'txtTglCetak.Text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub


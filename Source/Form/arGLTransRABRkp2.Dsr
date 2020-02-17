VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arGLTransRABRkp2 
   Caption         =   "Kwitansi"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "arGLTransRABRkp2.dsx":0000
End
Attribute VB_Name = "arGLTransRABRkp2"
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

End Sub

Private Sub Detail_Format()
With arGLTransRABRkp2.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.Text = sno
    'sItmCount = sItmCount + .Fields("jumlah").value
'    samtdebet = samtdebet + .Fields("amtdebet").value
'    samtkredit = samtkredit + .Fields("amtkredit").value
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_Format()
'txtpotonganfaktur.Text = sItmCount
'txttotdebet.Text = formatRupiah(samtdebet)
'txttotkredit.Text = formatRupiah(samtkredit)

End Sub

Private Sub GroupHeader2_AfterPrint()
sno = 0
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.Text & " dari " & Me.Pages.Count  'Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.Text = "Tanggal : " & Format(Now(), "DD-MM-YYYY HH:MM")
End Sub

'Private Sub PageHeader_BeforePrint()
'txtPeriode.Text = "Periode : " & arGLTrailBalance.adoKu.Recordset.Fields("yop").value & " - " & arGLTrailBalance.adoKu.Recordset.Fields("mop").value
'End Sub


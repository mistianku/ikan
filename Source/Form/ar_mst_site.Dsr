VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ar_mst_site 
   Caption         =   "Monitoring In Out Dokumen"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   330
   ClientWidth     =   20250
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "ar_mst_site.dsx":0000
End
Attribute VB_Name = "ar_mst_site"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim sttlsite As Double
Dim spage As Integer

Private Sub Detail_Format()
With ar_mst_site.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.Text = sno
    sttlsite = sttlsite + 1
    'sItmCount = sItmCount + .Fields("jumlah").value
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_Format()
sno = 0
End Sub

Private Sub GroupFooter2_BeforePrint()
lblttlsite = "Total Site"
txtttlsite = sttlsite
End Sub

Private Sub GroupHeader1_BeforePrint()
With ar_mst_site.adoKu.Recordset

lblheaderregion = "Region : " & .Fields("regionname").value & " # " & .Fields("regioncode").value
End With
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.Text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.Text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub




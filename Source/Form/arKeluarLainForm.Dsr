VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arKeluarLainForm 
   Caption         =   "Kwitansi"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   19080
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33655
   _ExtentY        =   19315
   SectionData     =   "arKeluarLainForm.dsx":0000
End
Attribute VB_Name = "arKeluarLainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double

Private Sub ActiveReport_ReportStart()
lblCompany1.Font.Size = 10
lblCompany2.Font.Size = 7
lblCompany2.Font.Size = 7

End Sub

Private Sub Detail_Format()
With arKeluarLainForm.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.Text = sno
    sItmCount = sItmCount + .Fields("jumlah").value
'    lblketerangan.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arKeluarLainForm.adoKu.Recordset.Fields("referensi").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_Format()
txtpotonganfaktur.Text = sItmCount

End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.Text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.Text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub

VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arMasukLainForm 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arMasukLainForm.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arMasukLainForm.dsx":628A
End
Attribute VB_Name = "arMasukLainForm"
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
With arMasukLainForm.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
    sItmCount = sItmCount + .Fields("jumlah").value
'    lblketerangan.Caption = ": " & arMasukLainForm.adoKu.Recordset.Fields("keterangan").value
'    lblreferensi.Caption = ": " & arMasukLainForm.adoKu.Recordset.Fields("referensi").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_Format()
txtpotonganfaktur.text = sItmCount

End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub

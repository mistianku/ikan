VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arlhppForm 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arlhppForm.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arlhppForm.dsx":628A
End
Attribute VB_Name = "arlhppForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim skolektor As Integer
Dim skodekolektorkey As String
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
        If skodekolektorkey <> .Fields("nodokumen") Then
           skodekolektorkey = .Fields("nodokumen")
           skolektor = skolektor + 1
        End If
    End If
End With
End Sub

Private Sub GroupFooter1_Format()
If skolektor = 1 Then
    GroupFooter1.Visible = False
Else
    GroupFooter1.Visible = True
End If
End Sub

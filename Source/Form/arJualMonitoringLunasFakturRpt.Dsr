VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arJualMonitoringLunasFakturRpt 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arJualMonitoringLunasFakturRpt.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arJualMonitoringLunasFakturRpt.dsx":628A
End
Attribute VB_Name = "arJualMonitoringLunasFakturRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double

Private Sub ActiveReport_ReportStart()
'lblCompany1.Font.Size = 10
'lblCompany2.Font.Size = 7
'lblCompany2.Font.Size = 7
lblCompany1.Caption = arJualMonitoringLunasFakturRpt.adoKu.Recordset.Fields("CmnyName").value
lblalamat.Caption = arJualMonitoringLunasFakturRpt.adoKu.Recordset.Fields("Address1").value
lblKota = arJualMonitoringLunasFakturRpt.adoKu.Recordset.Fields("City").value
lblFax = arJualMonitoringLunasFakturRpt.adoKu.Recordset.Fields("Faximale").value & " " & arJualMonitoringLunasFakturRpt.adoKu.Recordset.Fields("Phone1").value

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
With arJualMonitoringLunasFakturRpt.adoKu.Recordset

If Not .EOF Then

'    If Not .Fields("keysortby").value = 1 Then
'        arJualMonitoringLunasFakturRpt.Label26.Caption = ""
'        arJualMonitoringLunasFakturRpt.txtjumlah.Visible = False
'    Else
'        arJualMonitoringLunasFakturRpt.txtjumlah.Visible = True
'    End If
'    If .Fields("tukarfaktur").value = "Y" Then
'        txttgltfakturN.Visible = False
'        txttgltfakturY.Visible = True
'    Else
'        txttgltfakturN.Visible = True
'        txttgltfakturY.Visible = False
'    End If
    sno = sno + 1
    txtNo.text = sno

Else

End If
End With
End Sub


Private Sub GroupHeader1_BeforePrint()
sno = 0
End Sub

Private Sub GroupHeader1_Format()
With arJualMonitoringLunasFakturRpt.adoKu.Recordset
    
   If .Fields("keysortby").value = 0 Then
        arJualMonitoringLunasFakturRpt.GroupHeader1.Visible = False
        arJualMonitoringLunasFakturRpt.GroupFooter1.Visible = False
   Else
         arJualMonitoringLunasFakturRpt.GroupHeader1.Visible = True
        arJualMonitoringLunasFakturRpt.GroupFooter1.Visible = True
   End If
   
End With
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = Format(Now(), "DD-MM-YYYY HH:MM")
End Sub


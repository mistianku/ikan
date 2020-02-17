VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arGLTransGLRkp 
   Caption         =   "Kwitansi"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "arGLTransGLRkp.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26670
   _ExtentY        =   14870
   SectionData     =   "arGLTransGLRkp.dsx":628A
End
Attribute VB_Name = "arGLTransGLRkp"
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
    Image1.SizeMode = ddSMStretch
End If

If MenuFrm.sis_image = 0 Then
    Image1.Visible = False
    lblCompany1.Left = 90
    lblCompany2.Left = 90
    lblCompany3.Left = 90

End If

End Sub

Private Sub Detail_Format()
With Me.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
Else
End If
End With
End Sub

Private Sub GroupFooter1_AfterPrint()
sno = 0
End Sub

Private Sub GroupFooter1_BeforePrint()
Label62.Caption = "Grand.Total per " & Left(txttxtperiode.text, 4) & "-" & Right(txttxtperiode.text, 2) & " :"
End Sub

Private Sub GroupFooter3_BeforePrint()
Label60.Caption = "Sub.Total per " & txtnoslip.text & " :"
End Sub

Private Sub GroupHeader2_BeforePrint()
Label61.Caption = "Sub.Total per " & Field13.text & " :"
End Sub

Private Sub PageFooter_BeforePrint()
lblHalaman.Caption = "Halaman ke : " & txtPageNumber.text & " dari " & Me.Pages.Count  'Me.Pages.Count
lblUser.Caption = "User : " & MenuFrm.sUserID
txtTglCetak.text = "Tanggal : " & Format(Now(), "DD-MM-YYYY HH:MM")
End Sub



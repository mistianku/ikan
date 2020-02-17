VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arKwitansiForm 
   Caption         =   "Kwitansi"
   ClientHeight    =   11055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20370
   Icon            =   "arKwitansiForm.dsx":0000
   PaletteMode     =   2  'Custom
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "arKwitansiForm.dsx":628A
End
Attribute VB_Name = "arKwitansiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim terbilang As New CRUFLFungsiku.Konversi

Private Sub ActiveReport_ReportStart()

If arKwitansiForm.adoKu.Recordset.Fields("is_print_company").value = "1" Then
    lblCompany1.Visible = True
    Label37.Visible = True
    lblalamat.Visible = True
    lblKota.Visible = True
    lblFax.Visible = True
    Image1.Visible = True
Else
    lblCompany1.Visible = False
    Label37.Visible = False
    lblalamat.Visible = False
    lblKota.Visible = False
    lblFax.Visible = False
    Image1.Visible = False
End If

lblCompany1.Caption = arKwitansiForm.adoKu.Recordset.Fields("CmnyName").value
Label37.Caption = arKwitansiForm.adoKu.Recordset.Fields("PKPAddress2").value
lblalamat.Caption = arKwitansiForm.adoKu.Recordset.Fields("Address1").value
lblKota = arKwitansiForm.adoKu.Recordset.Fields("City").value
lblFax = arKwitansiForm.adoKu.Recordset.Fields("Faximale").value & " " & arKwitansiForm.adoKu.Recordset.Fields("Phone1").value

'If MenuFrm.simage_name = "" Then
'    foto = ""
'Else
'    foto = App.Path & "\images\" & MenuFrm.simage_name
'    Image1.Picture = LoadPicture(foto)
'End If

If arKwitansiForm.adoKu.Recordset.Fields("image_name").value = "" Then
    foto = ""
Else
    foto = App.Path & "\images\" & arKwitansiForm.adoKu.Recordset.Fields("image_name").value
    Image1.Picture = LoadPicture(foto)
End If

If arKwitansiForm.adoKu.Recordset.Fields("is_image").value = 0 Then
    Image1.Visible = False
    lblCompany1.Left = 360
    lblalamat.Left = 360
    lblKota.Left = 360
    lblFax.Left = 360
    Label37.Left = 360
End If

End Sub

Private Sub Detail_Format()
With arKwitansiForm.adoKu.Recordset

If Not .EOF Then
    sno = sno + 1
    'txtNo.Text = sno
    'sItmCount = sItmCount + .Fields("jumlah").value
Else
End If
End With
End Sub

Private Sub GroupFooter1_BeforePrint()
'lblterbilang.Caption = terbilang.terbilang(Field10)
End Sub

Private Sub GroupHeader1_Format()
With arKwitansiForm
If arKwitansiForm.adoKu.Recordset.Fields("CmpnyID").value <> "K001" Then
       
        .lblCompany1.Left = 360
        .Label37.Left = 360
        .lblalamat.Left = 360
        .lblKota.Left = 360
        .lblFax.Left = 360
End If
End With
End Sub


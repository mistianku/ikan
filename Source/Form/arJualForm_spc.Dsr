VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arJualForm_spc 
   Caption         =   "Kwitansi"
   ClientHeight    =   11055
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20370
   Icon            =   "arJualForm_spc.dsx":0000
   PaletteMode     =   2  'Custom
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "arJualForm_spc.dsx":628A
End
Attribute VB_Name = "arJualForm_spc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sno As Integer
Dim snodok As String
Dim sItmCount As Double
Dim foto As String


Private Sub ActiveReport_ReportStart()
If arJualForm_spc.adoKu.Recordset.Fields("is_print_company").value = "1" Then
    lblCompany1.Visible = True
    lblalamat.Visible = True
    Label37.Visible = True
    lblKota.Visible = True
    lblFax.Visible = True
    Image2.Visible = True
Else
    lblCompany1.Visible = False
    lblalamat.Visible = False
    Label37.Visible = False
    lblKota.Visible = False
    lblFax.Visible = False
    Image2.Visible = False
End If

Label40.Caption = arJualForm_spc.adoKu.Recordset.Fields("NPWP").value
Label42.Caption = arJualForm_spc.adoKu.Recordset.Fields("PKPName").value
Label45.Caption = arJualForm_spc.adoKu.Recordset.Fields("PKPAddress1").value
Label37.Caption = arJualForm_spc.adoKu.Recordset.Fields("PKPAddress2").value

If arJualForm_spc.adoKu.Recordset.Fields("NPWP").value = "" Then
    Label43.Caption = ""
Else
    Label43.Caption = "No. Rek : "
End If
lblCompany1.Caption = arJualForm_spc.adoKu.Recordset.Fields("CmnyName").value
lblalamat.Caption = arJualForm_spc.adoKu.Recordset.Fields("Address1").value
Label47.Caption = arJualForm_spc.adoKu.Recordset.Fields("alamat1").value
lblKota = arJualForm_spc.adoKu.Recordset.Fields("City").value
lblFax = arJualForm_spc.adoKu.Recordset.Fields("Faximale").value & " " & arJualForm_spc.adoKu.Recordset.Fields("Phone1").value


If arJualForm_spc.adoKu.Recordset.Fields("image_name").value = "" Then
    foto = ""
Else
    foto = App.Path & "\images\" & arJualForm_spc.adoKu.Recordset.Fields("image_name").value
    Image2.Picture = LoadPicture(foto)
End If
End Sub

Private Sub Detail_Format()
With arJualForm_spc.adoKu.Recordset


If Not .EOF Then
    sno = sno + 1
    txtNo.text = sno
    sItmCount = sItmCount + .Fields("jumlah").value
Else
End If
End With

End Sub


Private Sub GroupHeader1_Format()
With arJualForm_spc

If arJualForm_spc.adoKu.Recordset.Fields("is_image").value <> 1 Then
        arJualForm_spc.Image2.Visible = False
        .lblCompany1.Left = 360
        .Label37.Left = 360
        .lblalamat.Left = 360
        .lblKota.Left = 360
        .lblFax.Left = 360
End If
End With
End Sub



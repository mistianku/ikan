VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{5DC35748-D70A-417E-93B7-A488F085B02F}#90.0#0"; "smartnetbutton.ocx"
Begin VB.Form main_rpt 
   Caption         =   "Product Master"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   6
      Text            =   "Whs ID"
      Top             =   1800
      Width           =   5010
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   2415
      TabIndex        =   5
      Text            =   "Whs ID"
      Top             =   1800
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2415
      TabIndex        =   2
      Text            =   "Whs ID"
      Top             =   1440
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Text            =   "Whs ID"
      Top             =   1440
      Width           =   5010
   End
   Begin SmartNetButtonProject.SmartNetButton Browseku 
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   -2147483637
      Picture         =   "user_rpt2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAreaLayout=   2
   End
   Begin SmartNetButtonProject.SmartNetButton Browseku 
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BackColor       =   -2147483637
      Picture         =   "user_rpt2.frx":015A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionAreaLayout=   2
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   3120
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Pengguna Aplikasi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   8475
      Left            =   120
      Top             =   120
      Width           =   15
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   120
      Top             =   495
      Width           =   9555
   End
End
Attribute VB_Name = "main_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim istatus As StatusForm

Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub

Public Sub Execution()
On Error GoTo errhandler
Dim sActive As String
Dim sShortBy As Integer
Dim sDetail As Integer
Dim sShowHeadDetail As Integer

If Option1(0).value = True Then sActive = "3"
If Option1(1).value = True Then sActive = "1"
If Option1(2).value = True Then sActive = "0"

If Option2(0).value = True Then sShortBy = 1
If Option2(1).value = True Then sShortBy = 2
If Option2(2).value = True Then sShortBy = 3
If Option2(3).value = True Then sShortBy = 4

If Check2(0).value = 1 Then
    sShowHeadDetail = 1
Else
    sShowHeadDetail = 0
End If

Me.CR1.Reset
Me.CR1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.CR1.ReportFileName = App.Path + "\Reports\ProductMasterPriceRpt.Rpt"


Me.CR1.ParameterFields(0) = "@Priceid1" & ";" & Text1(19).Text & ";" & True
Me.CR1.ParameterFields(1) = "@Priceid2" & ";" & Text1(17).Text & ";" & True
Me.CR1.ParameterFields(2) = "@active" & ";" & sActive & ";" & True
Me.CR1.ParameterFields(3) = "@ProductType1" & ";" & Text1(0).Text & ";" & True
Me.CR1.ParameterFields(4) = "@ProductType2" & ";" & Text1(2).Text & ";" & True
Me.CR1.ParameterFields(5) = "@Groupid1" & ";" & Text1(4).Text & ";" & True
Me.CR1.ParameterFields(6) = "@Groupid2" & ";" & Text1(6).Text & ";" & True
Me.CR1.ParameterFields(7) = "@CatgryID1" & ";" & Text1(8).Text & ";" & True
Me.CR1.ParameterFields(8) = "@CatgryID2" & ";" & Text1(10).Text & ";" & True
Me.CR1.ParameterFields(9) = "@SubCatgryID1" & ";" & Text1(15).Text & ";" & True
Me.CR1.ParameterFields(10) = "@SubCatgryID2" & ";" & Text1(13).Text & ";" & True
Me.CR1.ParameterFields(11) = "@ProductID1" & ";" & Text1(20).Text & ";" & True
Me.CR1.ParameterFields(12) = "@ProductID2" & ";" & Text1(22).Text & ";" & True
Me.CR1.ParameterFields(13) = "@UserId" & ";" & MenuFrm.sUserID & ";" & True 'MenuFrm.sUserID
Me.CR1.ParameterFields(14) = "@SortBy" & ";" & sShortBy & ";" & True
Me.CR1.ParameterFields(15) = "@ShowHeadDetail" & ";" & sShowHeadDetail & ";" & True
Me.CR1.ParameterFields(16) = "@ByPrice" & ";" & 0 & ";" & True

Me.CR1.Destination = crptToWindow
Me.CR1.RetrieveDataFiles
Me.CR1.WindowState = crptMaximized
Me.CR1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Private Sub Browseku_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsProductType, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(0).Text = oBrowse.YangDipilih
        Text1(1).Text = oBrowse.Keterangan  'FindDataDetail(Text1(0), "PosProductType", "ProductType", "ProductTypeName", Parkir)
        Text1(0).SetFocus
    End If
Case 1
    oBrowse.ShowFinder BrowsProductType, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(2).Text = oBrowse.YangDipilih
        Text1(3).Text = oBrowse.Keterangan  'FindDataDetail(Text1(2), "PosProductType", "ProductType", "ProductTypeName", Parkir)
        Text1(2).SetFocus
    End If
Case 2
    oBrowse.ShowFinder BrowsGroup, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(4).Text = oBrowse.YangDipilih
        Text1(5).Text = oBrowse.Keterangan
        Text1(4).SetFocus
    End If
Case 3
    oBrowse.ShowFinder BrowsGroup, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(6).Text = oBrowse.YangDipilih
        Text1(7).Text = oBrowse.Keterangan
        Text1(6).SetFocus
    End If
Case 4
    oBrowse.ShowFinder BrowsCategory, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(8).Text = oBrowse.YangDipilih
        Text1(9).Text = oBrowse.Keterangan
        Text1(8).SetFocus
    End If
Case 5
    oBrowse.ShowFinder BrowsCategory, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(10).Text = oBrowse.YangDipilih
        Text1(11).Text = oBrowse.Keterangan
        Text1(10).SetFocus
    End If
Case 6
    oBrowse.ShowFinder BrowsSubCategory, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(15).Text = oBrowse.YangDipilih
        Text1(14).Text = oBrowse.Keterangan
        Text1(15).SetFocus
    End If
Case 7
    oBrowse.ShowFinder BrowsSubCategory, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(13).Text = oBrowse.YangDipilih
        Text1(12).Text = oBrowse.Keterangan
        Text1(13).SetFocus
    End If
Case 10
    oBrowse.ShowFinder BrowsItem, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(20).Text = oBrowse.YangDipilih
        Text1(21).Text = oBrowse.Keterangan
        Text1(20).SetFocus
    End If
Case 11
    oBrowse.ShowFinder BrowsItem, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(22).Text = oBrowse.YangDipilih
        Text1(23).Text = oBrowse.Keterangan
        Text1(22).SetFocus
    End If
Case 8
    oBrowse.ShowFinder BrowsDiscount, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(19).Text = oBrowse.YangDipilih
        Text1(18).Text = oBrowse.Keterangan
        Text1(19).SetFocus
    End If
Case 9
    oBrowse.ShowFinder BrowsDiscount, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(17).Text = oBrowse.YangDipilih
        Text1(16).Text = oBrowse.Keterangan
        Text1(17).SetFocus
    End If
End Select
Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Product Master Discount Report"
lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbar RefrshRpt
End Sub
Private Sub Form_Load()
cleardata
setPosisiBrowseKu Browseku(0), Text1(0)
setPosisiBrowseKu Browseku(1), Text1(2)
setPosisiBrowseKu Browseku(2), Text1(4)
setPosisiBrowseKu Browseku(3), Text1(6)
setPosisiBrowseKu Browseku(4), Text1(8)
setPosisiBrowseKu Browseku(5), Text1(10)
istatus = RefrshRpt
MenuFrm.SetToolbar istatus
End Sub
Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next

Text1(0) = oFindByQuery("Select top 1 ProductType from PosProductType Order by ProductType asc", Parkir)
Text1(1) = FindDataDetail(Text1(0), "PosProductType", "ProductType", "ProductTypeName", Parkir)
Text1(2) = oFindByQuery("Select top 1 ProductType from PosProductType Order by ProductType Desc", Parkir)
Text1(3) = FindDataDetail(Text1(2), "PosProductType", "ProductType", "ProductTypeName", Parkir)

Text1(4) = oFindByQuery("Select top 1 GroupID from PosGroup Order by GroupID asc", Parkir)
Text1(5) = FindDataDetail(Text1(4), "PosGroup", "GroupID", "GroupName", Parkir)
Text1(6) = oFindByQuery("Select top 1 GroupID from PosGroup Order by GroupID Desc", Parkir)
Text1(7) = FindDataDetail(Text1(6), "PosGroup", "GroupID", "GroupName", Parkir)

Text1(8) = oFindByQuery("Select top 1 CatgryID from PosCategory Order by CatgryID asc", Parkir)
Text1(9) = FindDataDetail(Text1(8), "PosCategory", "CatgryID", "CategoryName", Parkir)
Text1(10) = oFindByQuery("Select top 1 CatgryID from PosCategory Order by CatgryID Desc", Parkir)
Text1(11) = FindDataDetail(Text1(10), "PosCategory", "CatgryID", "CategoryName", Parkir)

Text1(15) = oFindByQuery("Select top 1 SubCatgryID from PosSubCategory Order by SubCatgryID asc", Parkir)
Text1(14) = FindDataDetail(Text1(15), "PosSubCategory", "SubCatgryID", "SubCatgryName", Parkir)
Text1(13) = oFindByQuery("Select top 1 SubCatgryID from PosSubCategory Order by SubCatgryID Desc", Parkir)
Text1(12) = FindDataDetail(Text1(13), "PosSubCategory", "SubCatgryID", "SubCatgryName", Parkir)

Text1(20) = oFindByQuery("Select top 1 ProductID from PosItem Order by ProductID asc", Parkir)
Text1(21) = FindDataDetail(Text1(20), "PosItem", "ProductID", "ProductName", Parkir)
Text1(22) = oFindByQuery("Select top 1 ProductID from PosItem Order by ProductID Desc", Parkir)
Text1(23) = FindDataDetail(Text1(22), "PosItem", "ProductID", "ProductName", Parkir)

Text1(19) = oFindByQuery("Select top 1 DiscCode from PosDiscount Order by DiscCode asc", Parkir)
Text1(18) = FindDataDetail(Text1(19), "PosDiscount", "DiscCode", "DiscName", Parkir)
Text1(17) = oFindByQuery("Select top 1 DiscCode from PosDiscount Order by DiscCode Desc", Parkir)
Text1(16) = FindDataDetail(Text1(17), "PosDiscount", "DiscCode", "DiscName", Parkir)

End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
Case 1
Case 2
Case 3
End Select
End Sub

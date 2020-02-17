VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form MutasiStockRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Mutasi Stok Report"
   ClientHeight    =   5835
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   10170
   ControlBox      =   0   'False
   DrawMode        =   7  'Invert
   DrawStyle       =   6  'Inside Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Sort By"
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   10455
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   4800
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2340
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Width           =   5535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fungsi"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Kategori"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Brand"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   11
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MutasiStockRpt.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   15
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MutasiStockRpt.frx":001C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai Dengan"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kartu Stock "
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10455
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   18
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "MutasiStockRpt.frx":0038
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   1935
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   4
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MutasiStockRpt.frx":0054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   19
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "MutasiStockRpt.frx":0070
      End
      Begin VB.Label Label1 
         Caption         =   "Periode"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Warehouse"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "MutasiStockRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0

    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    If Option1(0).value = True Then
        oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    Else
    If Option1(1).value = True Then
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
        Else
        oBrowse.ShowFinder BrowsFunction, "", ubAscending, DBaseConection.Modul
    End If
    End If
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
Case 2
    If Option1(0).value = True Then
        oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    Else
    If Option1(1).value = True Then
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
        Else
        oBrowse.ShowFinder BrowsFunction, "", ubAscending, DBaseConection.Modul
    End If
    End If
    If Not oBrowse.YangDipilih = "" Then
        Text1(4) = oBrowse.YangDipilih
        Text1(5) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Mutasi Stok Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMutasiStockRpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatFrameBackground Frame1(0)
oFormatOption 1, Me
istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width

MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"
oGetComboBoxTahun FlatComboBox1(0)
oGetComboBulanan FlatComboBox1(1)

Text1(1) = oFindByQuery("Select kodebrand from master_brand order by kodebrand asc limit 1 ", DBaseConection.Modul)
Text1(3) = oFindByQuery("Select namabrand from master_brand order by kodebrand asc limit 1 ", DBaseConection.Modul)
Text1(4) = oFindByQuery("Select kodebrand from master_brand order by kodebrand desc limit 1 ", DBaseConection.Modul)
Text1(5) = oFindByQuery("Select namabrand from master_brand order by kodebrand desc limit 1 ", DBaseConection.Modul)

End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub
Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim syop As Integer
Dim smop, sget, scek As Integer
Dim skodefr As String
Dim skodeto As String
Dim txtmessage, txtfilterby As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

syop = FlatComboBox1(0).text
smop = FlatComboBox1(1).ListIndex + 1
sget = IIf(Option1(0).value = True, 1, IIf(Option1(1).value = True, 2, 3))
txtfilterby = "Filter By " & IIf(Option1(0).value = True, "Brand", IIf(Option1(1).value = True, "Kategori", "Fungsi"))
sQuery = "CALL sp_mutasi_stock_rpt('"
sQuery = sQuery & syop & "','"
sQuery = sQuery & smop & "','"
sQuery = sQuery & Text1(0) & "','"
sQuery = sQuery & Text1(1) & "','"
sQuery = sQuery & Text1(4) & "','"
sQuery = sQuery & sget & "',"

'sp_mutasi_stock_rpt`(IN syop INT,smop INT, skodegudang CHAR(6),
    'skodebrandfr CHAR(6),skodebrandto CHAR(6),sget INT,scek INT)
    
If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If
With arMutasiStockFrm
    .lblHeaderTrx = "Mutasi Stock"
    .lblFilter.Caption = txtfilterby
    .txtPeriode = smop & "-" & syop
    .txtFilter.text = Text1(1) & "-" & Text1(3) & " S/d " & Text1(4) & "-" & Text1(5)
    
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery & "0)"
    

 '   .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
    .Show
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Mutasi Stock"
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Text1(1) = oFindByQuery("Select kodebrand from master_brand order by kodebrand asc limit 1 ", DBaseConection.Modul)
    Text1(3) = oFindByQuery("Select namabrand from master_brand order by kodebrand asc limit 1 ", DBaseConection.Modul)
    Text1(4) = oFindByQuery("Select kodebrand from master_brand order by kodebrand desc limit 1 ", DBaseConection.Modul)
    Text1(5) = oFindByQuery("Select namabrand from master_brand order by kodebrand desc limit 1 ", DBaseConection.Modul)
Case 1
    Text1(1) = oFindByQuery("Select kodekategori from master_kategori order by kodekategori asc limit 1 ", DBaseConection.Modul)
    Text1(3) = oFindByQuery("Select namakategori from master_kategori order by kodekategori asc limit 1 ", DBaseConection.Modul)
    Text1(4) = oFindByQuery("Select kodekategori from master_kategori order by kodekategori desc limit 1 ", DBaseConection.Modul)
    Text1(5) = oFindByQuery("Select namakategori from master_kategori order by kodekategori desc limit 1 ", DBaseConection.Modul)
Case 2
    Text1(1) = oFindByQuery("Select kodefungsi from master_fungsi order by kodefungsi asc limit 1 ", DBaseConection.Modul)
    Text1(3) = oFindByQuery("Select namafungsi from master_fungsi order by kodefungsi asc limit 1 ", DBaseConection.Modul)
    Text1(4) = oFindByQuery("Select kodefungsi from master_fungsi order by kodefungsi desc limit 1 ", DBaseConection.Modul)
    Text1(5) = oFindByQuery("Select namafungsi from master_fungsi order by kodefungsi desc limit 1 ", DBaseConection.Modul)
End Select
End Sub
Public Function sp_create_master_mutasi_stok_temp(saudituser As String, syop As Integer, smop As Integer, skodegudang As String, ssortby As String, ssortbyfr As String, ssortbyto As String)
On Error GoTo errhandler
Dim oConku As New ADODB.Connection
Dim oRstku As New ADODB.Recordset
sp_create_master_mutasi_stok_temp = False
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.Modul)
sQuery = "CALL sp_create_master_mutasi_stok_temp('"
sQuery = sQuery & saudituser & "'," & syop & "," & smop & ",'" & skodegudang & "','"
sQuery = sQuery & ssortby & "','" & ssortbyfr & "','" & ssortbyto & "')"
oConku.Execute (sQuery)
sp_create_master_mutasi_stok_temp = True
oConku.Close

    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetProductID"
End Function

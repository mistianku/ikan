VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form BrowseFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daftar Data"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Pencarian Otomatis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      MouseIcon       =   "BrowseFrm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pilih"
   End
   Begin VSFlex8LCtl.VSFlexGrid GridBrowse 
      Height          =   3615
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   12135
      _cx             =   21405
      _cy             =   6376
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   5
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSDFLATS.FlatComboBox fCombo1 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      MouseIcon       =   "BrowseFrm.frx":001C
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   10275
   End
   Begin VSDFLATS.FlatComboBox fCombo1 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      MouseIcon       =   "BrowseFrm.frx":0038
   End
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      MouseIcon       =   "BrowseFrm.frx":0054
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Batal"
   End
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Index           =   2
      Left            =   9960
      TabIndex        =   10
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      MouseIcon       =   "BrowseFrm.frx":0070
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cari"
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Kata Pencarian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Degan Kriteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Cari Berdasarkan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "BrowseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sBrowsCriteria As String
Dim svalue As String
Dim sDesc As String
Dim sBrowse As BrowseTables
Dim sDBaseConectionQ As DBaseConection
Dim sTable As String
Dim sField As String
Dim sFields As String
Dim sFieldDefault As Double
Dim sOrderBy As urutBy
Dim sListIndex As Double
Public Sub ShowFinder(WhatToBrowse As BrowseTables, sCriteriaku As String, sUrutBy As urutBy, sDBaseConection As DBaseConection)
    sBrowse = WhatToBrowse
    sOrderBy = sUrutBy
    SetForm sBrowse
    sBrowsCriteria = sCriteriaku
    sDBaseConectionQ = sDBaseConection
    InitData sBrowse, sCriteriaku, sOrderBy, sDBaseConectionQ
    Me.Show 1
End Sub
Private Sub InitData(WhatToBrowse As BrowseTables, sCriteriaku As String, sUrutBy As urutBy, sDBaseConection As DBaseConection)
On Error GoTo errhandler
    Dim oCon As New ADODB.Connection
    Dim oRs As New ADODB.Recordset
    Dim sQuery As String, sKriteria As String, sField As String, sKriteria2 As String
    If oCon.State = 1 Then oCon.Close

    oCon.Open MainModule.Conectionku(sDBaseConection)
    ClearGrid
    sKriteria = ToText(Text1.text)
    sField = ToText(fCombo1(0).text)
    If sCriteriaku = "" Then
        sKriteria2 = "''=''"
    Else
        sKriteria2 = sCriteriaku
    End If
    If sKriteria <> "" Then
        If fCombo1(1).ListIndex = sListIndex Then
            sKriteria = "Like '" & sKriteria & "%'"
        Else
            sKriteria = "Like '%" & sKriteria & "%'"
        End If
        
        sKriteria = " Where " & sField & " " & sKriteria & " And " & sKriteria2
    Else
        sKriteria = " Where " & sKriteria2
    End If
    sQuery = "Select " & sFields & " from " & sTable & " " & sKriteria & " order by " & fCombo1(0).List(0) & IIf(sUrutBy = ubAscending, " asc", " desc") & " limit 7000"
    Set oRs = oCon.Execute(sQuery)
    With GridBrowse
    GridBrowse.ColAlignment(0) = flexAlignRightCenter
        Dim irow As Integer, i As Integer
        Do While Not oRs.EOF
            .Rows = .Rows + 1
            irow = irow + 1
            For i = 0 To oRs.Fields.Count - 1
                .TextMatrix(irow, i) = IIf(IsNull(oRs(i)), "", oRs(i))
            Next
            oRs.MoveNext
        Loop
    End With
    If sListIndex = 1 Then
        fCombo1(0).ListIndex = sFieldDefault
        sListIndex = 0
    End If
Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "InitData"
End Sub

Private Sub SetForm(WhatToBrowse As BrowseTables)
Dim LebarBrowsing As Integer
LebarBrowsing = GridBrowse.Width
fCombo1(1).AddItem "Mempunyai Awal"
fCombo1(1).AddItem "Berisikan"
Select Case WhatToBrowse

Case BrowseTables.BrowsUser
    AddToCombo "UserID"
    AddToCombo "NamaUser"
    ClearGrid
    '12975
    AddColom "Kode User", 1500, ratakanan
    AddColom "Nama User", LebarBrowsing - (1500), ratakiri
    sTable = "master_User"
    sFieldDefault = 1
Case BrowseTables.BrowsLogin
    AddToCombo "UserID"
    AddToCombo "NamaUser"
    ClearGrid
    AddColom "Kode User", 1500, ratakanan
    AddColom "Nama User", LebarBrowsing - 1500, ratakiri
    sTable = "master_User"
    sFieldDefault = 1
Case BrowseTables.BrowsAgama
    AddToCombo "kode"
    AddToCombo "agama"
    ClearGrid
    AddColom "Kode", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "master_agama"
    sFieldDefault = 1
Case BrowseTables.BrowsPekerjaan
    AddToCombo "kode"
    AddToCombo "pekerjaan"
    ClearGrid
    AddColom "Kode", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "master_pekerjaan"
    sFieldDefault = 1
Case BrowseTables.BrowsCompany
    AddToCombo "cmpnyid"
    AddToCombo "CmpnyName"
    
    ClearGrid
    AddColom "Kode Company", 1500, ratakanan
    AddColom "Nama ", LebarBrowsing - 3500, ratakiri
    AddColom "Alamat", 2000, ratakiri
    sTable = "company"
    sFieldDefault = 1
    
Case BrowseTables.BrowsMasterPreferencesSpecial
    AddToCombo "CmpnyID"
    AddToCombo "CmnyName"
    AddToCombo "Address1"
    ClearGrid
    AddColom "Kode Company", 1500, ratakanan
    AddColom "Nama ", LebarBrowsing - 3500, ratakiri
    AddColom "Alamat", 2000, ratakiri
    sTable = "master_preferences_special"
    sFieldDefault = 1
    
Case BrowseTables.BrowsPelajaranGroup
    AddToCombo "kodegroup"
    AddToCombo "namagroup"
    ClearGrid
    AddColom "Kode Group", 1500, ratakanan
    AddColom "Nama Group", LebarBrowsing - 1500, ratakiri
    sTable = "master_pelajaran_group"
    sFieldDefault = 1
    
Case BrowseTables.BrowsArea
    AddToCombo "kodearea"
    AddToCombo "namaarea"
    ClearGrid
    AddColom "Kode Area", 1500, ratakanan
    AddColom "Nama Area", LebarBrowsing - 1500, ratakiri
    sTable = "master_area"
    sFieldDefault = 1
    
Case BrowseTables.BrowsSettingDocument
    AddToCombo "docid"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "ID Dokumen", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - (1500), ratakiri
    sTable = "setting_document"
    sFieldDefault = 1
    
Case BrowseTables.BrowsPendaftaran
    AddToCombo "nopendaftaran"
    AddToCombo "tglpendaftaran"
    AddToCombo "noidsiswa"
    AddToCombo "nmlengkap"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Pendaftaran", 1500, ratakanan
    AddColom "Tgl Pendaftaran", 3000, ratatengah
    AddColom "No.Id Siswa", 1500, ratakiri
    AddColom "Nama Siswa", 3000, ratakiri
    AddColom "Keterangan ", 3000, ratakiri
    sTable = "vtransaksi_pendaftaran"
    sFieldDefault = 0
    
Case BrowseTables.BrowsPendaftaran
    AddToCombo "nopendaftaran"
    AddToCombo "tglpendaftaran"
    AddToCombo "noidsiswa"
    AddToCombo "nmlengkap"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Pendaftaran", 1500, ratakanan
    AddColom "Tgl Pendaftaran", 3000, ratatengah
    AddColom "No.Id Siswa", 1500, ratakiri
    AddColom "Nama Siswa", 3000, ratakiri
    AddColom "Keterangan ", 3000, ratakiri
    sTable = "vtransaksi_pendaftaran"
    sFieldDefault = 0
    
Case BrowseTables.BrowsKwitansi
    AddToCombo "nokwitansi"
    AddToCombo "tglkwitansi"
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    AddToCombo "almtrumah1"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No.Kwitansi", 1500, ratakanan
    AddColom "Tanggal", 3000, ratatengah
    AddColom "Kode Customer", 1500, ratakiri
    AddColom "Nama Customer", 3000, ratakiri
    AddColom "Alamat", 3000, ratakiri
    AddColom "Keterangan ", 3000, ratakiri
    sTable = "vtransaksi_kwitansi"
    sFieldDefault = 0
   
'Case BrowseTables.BrowsKwitansi
'    AddToCombo "nokwitansi"
'    AddToCombo "tglkwitansi"
'    AddToCombo "noidsiswa"
'    AddToCombo "nmlengkap"
'    AddToCombo "almtrumah1"
'    AddToCombo "keterangan"
'    ClearGrid
'    AddColom "No.Kwitansi", 1500, ratakanan
'    AddColom "Tanggal", 1000, ratatengah
'    AddColom "No.Id Siswa", 1000, ratakiri
'    AddColom "Nama Siswa", 3000, ratakiri
'    AddColom "Alamat", 3000, ratakiri
'    AddColom "Keterangan ", 3000, ratakiri
'    sTable = "vtransaksi_kwitansi"
    
Case BrowseTables.BrowsKelas
    AddToCombo "nokursus"
    AddToCombo "tglmulai"
    AddToCombo "noidsiswa"
    AddToCombo "nmlengkap"
    AddToCombo "nmpelajaran"
    AddToCombo "dokstatus"
    ClearGrid
    AddColom "No Kursus", 1500, ratakanan
    AddColom "Tgl Mulai", 1000, ratatengah
    AddColom "No.Id Siswa", 1000, ratakiri
    AddColom "Nama Siswa", 3000, ratakiri
    AddColom "Pelajaran ", 3000, ratakiri
    AddColom "Kartu Kelas", 1000, ratakiri
    sTable = "vmaster_kelas_brows"
    sFieldDefault = 0
'Case BrowseTables.BrowsPembimbing
'    AddToCombo "id"
'    AddToCombo "namapembimbing"
'    ClearGrid
'    AddColom "Id", 500, ratakanan
'    AddColom "Nama Pembimbing", LebarBrowsing - (1500), ratakiri
'    sTable = "master_pembimbing"

Case BrowseTables.BrowsKartuKelas
    AddToCombo "nokursus"
    AddToCombo "tglmulai"
    AddToCombo "keterangan"
    AddToCombo "kodelevel"
    ClearGrid
    AddColom "No Kursus", 1500, ratakanan
    AddColom "Tgl Mulai", 1000, ratatengah
    AddColom "Kelas", 5000, ratakiri
    AddColom "Level", 3000, ratakiri
    sTable = "vmaster_kartu_materi_kelas"
    sFieldDefault = 0
    
Case BrowseTables.BrowsCuti
    AddToCombo "nodokumen"
    AddToCombo "tanggal"
    AddToCombo "noidsiswa"
    AddToCombo "nmlengkap"
    AddToCombo "alamat"
    ClearGrid
    AddColom "No.Dok", 1500, ratakanan
    AddColom "Tanggal", 3000, ratatengah
    AddColom "No.Id Siswa", 1500, ratakiri
    AddColom "Nama Siswa", 3000, ratakiri
    AddColom "Alamat ", 3000, ratakiri
    sTable = "vtransaksi_cuti"
    sFieldDefault = 0
    
Case BrowseTables.BrowsBrand
    AddToCombo "kodebrand"
    AddToCombo "namabrand"
    ClearGrid
    AddColom "Kode Brand", 1500, ratakanan
    AddColom "Nama Brand", LebarBrowsing - 1500, ratakiri
    sTable = "master_brand"
    sFieldDefault = 1
    
Case BrowseTables.BrowsCategory
    AddToCombo "kodekategori"
    AddToCombo "namakategori"
    ClearGrid
    AddColom "Kode Kategori", 1500, ratakanan
    AddColom "Nama Kategori", LebarBrowsing - 1500, ratakiri
    sTable = "master_kategori"
    sFieldDefault = 1
    
Case BrowseTables.BrowsFunction
    AddToCombo "kodefungsi"
    AddToCombo "namafungsi"
    ClearGrid
    AddColom "Kode Fungsi", 1500, ratakanan
    AddColom "Nama Fungsi", LebarBrowsing - 1500, ratakiri
    sTable = "master_fungsi"
    sFieldDefault = 1
    
Case BrowseTables.BrowsHarga
    AddToCombo "kodeharga"
    AddToCombo "namaharga"
    ClearGrid
    AddColom "Kode Harga", 1500, ratakanan
    AddColom "Nama Harga", LebarBrowsing - 1500, ratakiri
    sTable = "master_harga"
    sFieldDefault = 1
    
Case BrowseTables.BrowsDiskon
    AddToCombo "kodediskon"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "Kode Diskon", 1500, ratakanan
    AddColom "Nama keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "master_diskon"
    sFieldDefault = 1
    
Case BrowseTables.BrowsFee
    AddToCombo "kodediskon"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "Kode Diskon", 1500, ratakanan
    AddColom "Nama keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "master_fee"
    sFieldDefault = 1
    
Case BrowseTables.BrowsTipeGudang
    AddToCombo "tipegudang"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "Tipe Gudang", 1500, ratakanan
    AddColom "Nama keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "master_tipegudang"
    sFieldDefault = 1
    
Case BrowseTables.BrowsMasterProduk
    AddToCombo "kodeproduk"
    AddToCombo "namaproduk"
    AddToCombo "stok"
    ClearGrid
    AddColom "Kode Produk", 1500, ratakanan
    AddColom "Nama Produk", LebarBrowsing - 3000, ratakiri
    AddColom "Jumlah Stock", 1500, ratakiri
    sTable = "vmaster_produk_brows"
    sFieldDefault = 1
    
    
Case BrowseTables.BrowsSatuanProduk
    AddToCombo "kodesatuan"
    AddToCombo "namasatuan"
    ClearGrid
    AddColom "Kode Satuan", 1500, ratakanan
    AddColom "Nama Satuan", LebarBrowsing - 1500, ratakiri
    sTable = "master_satuan_produk"
    sFieldDefault = 1
    
Case BrowseTables.BrowsGudang
    AddToCombo "kodegudang"
    AddToCombo "namagudang"
    AddToCombo "namatipegudang"
    ClearGrid
    AddColom "Kode Gudang", 1500, ratakanan
    AddColom "Nama Gudang", LebarBrowsing - 3500, ratakiri
    AddColom "Tipe Gudang", 2000, ratakiri
    sTable = "vmaster_gudang"
    sFieldDefault = 1
    
Case BrowseTables.BrowsSupplier
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    ClearGrid
    '12975
    AddColom "Kode Customer", 1500, ratakanan
    AddColom "Nama Customer", LebarBrowsing - (1500), ratakiri
    sTable = "master_supplier"
    sFieldDefault = 1
    
Case BrowseTables.Browscustomer
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    ClearGrid
    '12975
    AddColom "Kode Customer", 1500, ratakanan
    AddColom "Nama Customer", LebarBrowsing - (1500), ratakiri
    sTable = "master_customer"
    sFieldDefault = 1
    
Case BrowseTables.BrowsPembelian
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Dokumen", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Kode Customer", 1500, ratakiri
    AddColom "Nama Customer", (LebarBrowsing - 4500) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 4500) / 2, ratakiri
    sTable = "vtransaksi_masuk"
    sFieldDefault = 0
    
Case BrowseTables.BrowsMasukLain
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "kodesupplier"
    AddToCombo "namasupplier"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Dokumen", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Kode Supplier", 1500, ratakiri
    AddColom "Nama Supplier", (LebarBrowsing - 4500) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 4500) / 2, ratakiri
    sTable = "vtransaksi_masuk_lain"
    sFieldDefault = 0
    
Case BrowseTables.BrowsPenjualan
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Dokumen", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Kode Customer", 1500, ratakiri
    AddColom "Nama Customer", (LebarBrowsing - 4500) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 4500) / 2, ratakiri
    sTable = "vtransaksi_keluar"
    sFieldDefault = 0
    
Case BrowseTables.BrowsKeluarLain
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "kodecustomer"
    AddToCombo "namacustomer"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Dokumen", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Kode Customer", 1500, ratakiri
    AddColom "Nama Customer", (LebarBrowsing - 4500) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 4500) / 2, ratakiri
    sTable = "vtransaksi_keluar_lain"
    sFieldDefault = 0
    
Case BrowseTables.BrowsUserGroup
    AddToCombo "kodegroup"
    AddToCombo "groupuser"
    ClearGrid
    AddColom "Kode Group", 1500, ratakanan
    AddColom "Group User", LebarBrowsing - 1500, ratakiri
    sTable = "master_group_user"
    sFieldDefault = 1
    
Case BrowseTables.BrowsModule
    AddToCombo "Modulid"
    AddToCombo "ModuleMenu"
    AddToCombo "Dscription"
    ClearGrid
    AddColom "Modul ID", 1500, ratakanan
    AddColom "Modul Menu", 2000, ratakiri
    AddColom "Keterangan", LebarBrowsing - 3500, ratakiri
    sTable = "master_module"
    sFieldDefault = 1
    
Case BrowseTables.BrowsSiswaKeluar
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "noidsiswa"
    AddToCombo "nmlengkap"
    AddToCombo "alamat"
    ClearGrid
    AddColom "No.Dok", 1500, ratakanan
    AddColom "Tanggal", 3000, ratatengah
    AddColom "No.Id Siswa", 1500, ratakiri
    AddColom "Nama Siswa", 3000, ratakiri
    AddColom "Alamat ", 3000, ratakiri
    sTable = "vtransaksi_siswa_keluar"
    sFieldDefault = 0

Case BrowseTables.BrowsJenisMateri
    AddToCombo "kodejenis"
    AddToCombo "namajenis"
    ClearGrid
    AddColom "Kode Jenis", 1500, ratakanan
    AddColom "Nama Jenis", LebarBrowsing - 1500, ratakiri
    sTable = "master_jenis"
    sFieldDefault = 1
    
Case BrowseTables.BrowsTransfer
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "kodegudangfr"
    AddToCombo "kodegudangto"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Dokumen", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Dr Gudang", 1500, ratakiri
    AddColom "Ke Gudang", 1500, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 6000), ratakiri
    sTable = "vtransaksi_transfer"
    sFieldDefault = 0
    
Case BrowseTables.BrowsMasterDefaultPelajaran
    AddToCombo "defaultid"
    AddToCombo "keterangan"
    AddToCombo "pelajaran"
    AddToCombo "kodegroup"
    AddToCombo "namagroup"
    ClearGrid
    AddColom "Default ID", 1500, ratakanan
    AddColom "Materi Pelajaran", 1500, ratatengah
    AddColom "Kode Pelajaran", 1500, ratakiri
    AddColom "Kode Group", 1500, ratakiri
    AddColom "Nama Group", (LebarBrowsing - 6000), ratakiri
    sTable = "vmaster_default_pelajaran"
    sFieldDefault = 1
    
Case BrowseTables.BrowsAkunSumberData
    AddToCombo "kd_dataentry"
    AddToCombo "nm_dataentry"
    ClearGrid
    AddColom "Kode Entry Data", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - (1500), ratakiri
    sTable = "tblglsbrdata"
    sFieldDefault = 1
    
Case BrowseTables.BrowsAkunMasterCOA
    AddToCombo "coa"
    AddToCombo "nm_akun"
    ClearGrid
    AddColom "Kode Akun", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - (1500), ratakiri
    sTable = "tblglmas"
    sFieldDefault = 1
    
Case BrowseTables.BrowsAkunGroupSumberData
    AddToCombo "gr_dataentry"
    AddToCombo "nm_grupdata"
    ClearGrid
    AddColom "Group Entry", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - (1500), ratakiri
    sTable = "tblgrupdataentry"
    sFieldDefault = 1
    
Case BrowseTables.BrowsAkunRAB
    AddToCombo "kd_rab"
    AddToCombo "nm_rab"
    ClearGrid
    AddColom "Kode RAB", 1500, ratakanan
    AddColom "Nama RAB", LebarBrowsing - (1500), ratakiri
    sTable = "tblrab"
    sFieldDefault = 1
    
Case BrowseTables.BrowsAkunTransRAB
    AddToCombo "noslip"
    AddToCombo "tanggal"
    AddToCombo "referensi"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No RAB", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Referensi", (LebarBrowsing - 3000) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 3000) / 2, ratakiri
    sTable = "trnent_rab"
    sFieldDefault = 0
    
Case BrowseTables.BrowsAkunTransGL
    AddToCombo "notran"
    AddToCombo "tanggal"
    AddToCombo "referensi"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No Transaksi", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Referensi", (LebarBrowsing - 3000) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 3000) / 2, ratakiri
    sTable = "trnent_gl"
    sFieldDefault = 0
    
Case BrowseTables.BrowsSalesman
    AddToCombo "kodesalesman"
    AddToCombo "namasalesman"
    ClearGrid
    AddColom "Kode Salesman", 1500, ratakanan
    AddColom "Nama Salesman", LebarBrowsing - 1500, ratakiri
    sTable = "master_salesman"
    sFieldDefault = 1
    
Case BrowseTables.BrowsKolektor
    AddToCombo "kodekolektor"
    AddToCombo "namakolektor"
    ClearGrid
    AddColom "Kode Kolektor", 1500, ratakanan
    AddColom "Nama Kolektor", LebarBrowsing - 1500, ratakiri
    sTable = "master_kolektor"
    sFieldDefault = 1
    
Case BrowseTables.BrowsBarSettLabelBarcode
    AddToCombo "modellabel"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "Setting ID", 1500, ratakanan
    AddColom "Keterangan", LebarBrowsing - 1500, ratakiri
    sTable = "bar_sett_label_barcode"
    sFieldDefault = 1
    
Case BrowseTables.Browslhpp
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "referensi"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No. LHPP", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Referensi", (LebarBrowsing - 3000) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 3000) / 2, ratakiri
    sTable = "transaksi_lhpp"
    sFieldDefault = 0

Case BrowseTables.Browslhppdetail1
    AddToCombo "docentry"
    AddToCombo "kodecustomer"
    AddToCombo "jmlkwitansi"
    AddToCombo "totnilkwitansi"
    ClearGrid
    AddColom "Doc.Entry", 1500, ratakanan
    AddColom "Kode Customer", 1500, ratatengah
    AddColom "jumlah Kwitansi", 2000, ratakanan
    AddColom "Total Nil Kwitansi", 2000, ratakanan
    sTable = "transaksi_lhpp_detail1"
    sFieldDefault = 1
    
    
Case BrowseTables.Browslhppentry
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "referensi"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No. LHPP", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Referensi", (LebarBrowsing - 3000) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 3000) / 2, ratakiri
    sTable = "transaksi_lhpp_entry"
    sFieldDefault = 0
    
Case BrowseTables.Browslhppentrydetail1
    AddToCombo "docentry"
    AddToCombo "kodecustomer"
    AddToCombo "jmlkwitansi"
    AddToCombo "totnilkwitansi"
    ClearGrid
    AddColom "Doc.Entry", 1500, ratakanan
    AddColom "Kode Customer", 1500, ratatengah
    AddColom "jumlah Kwitansi", 2000, ratakanan
    AddColom "Total Nil Kwitansi", 2000, ratakanan
    sTable = "transaksi_lhpp_entry_detail1"
    sFieldDefault = 1
    
Case BrowseTables.Browslhpptf
    AddToCombo "nodokumen"
    AddToCombo "tgldokumen"
    AddToCombo "referensi"
    AddToCombo "keterangan"
    ClearGrid
    AddColom "No. LHPP", 1500, ratakanan
    AddColom "Tanggal", 1500, ratatengah
    AddColom "Referensi", (LebarBrowsing - 3000) / 2, ratakiri
    AddColom "Keterangan", (LebarBrowsing - 3000) / 2, ratakiri
    sTable = "transaksi_lhpp_tf"
    sFieldDefault = 0

Case BrowseTables.Browslhpptfdetail1
    AddToCombo "docentry"
    AddToCombo "kodecustomer"
    AddToCombo "jmlkwitansi"
    AddToCombo "totnilkwitansi"
    ClearGrid
    AddColom "Doc.Entry", 1500, ratakanan
    AddColom "Kode Customer", 1500, ratatengah
    AddColom "jumlah Kwitansi", 2000, ratakanan
    AddColom "Total Nil Kwitansi", 2000, ratakanan
    sTable = "transaksi_lhpp_tf_detail1"
    sFieldDefault = 0
    
Case BrowseTables.BrowsCompanyLogin
    AddToCombo "CmpnyId"
    AddToCombo "CmpnyName"
    ClearGrid
    AddColom "Kode Unit", 1500, ratakanan
    AddColom "Nama Unit", LebarBrowsing - 1500, ratakiri
    sTable = " v_company_login_browse"
    sFieldDefault = 1
    
'Case BrowseTables.BrowsGoodsAdjustmen
'    AddToCombo "Docnum"
'    AddToCombo "DocDate"
'    AddToCombo "DocStatus"
'    AddToCombo "Refference"
'    AddToCombo "Note"
'    AddToCombo "TotalAmmnt"
'    ClearGrid
'    AddColom "Doc.No", 1500
'    AddColom "Date", 1000
'    AddColom "Status", 1000
'    AddColom "Refference", 1500
'    AddColom "Note", 3000
'    AddColom "Ammount", 3000
'
'    sTable = "vBrowseAdjustmen"
'
'Case BrowseTables.BrowsSettingDocNumber
'    AddToCombo "DocId"
'    AddToCombo "Dscription"
'    ClearGrid
'    AddColom "Modul", 1500
'    AddColom "Description", LebarBrowsing - 1500
'    AddColom "Status", 750
'    sTable = "PosSettingDocNumber"
'
''Case BrowseTables.BrowsStockOpname
''            sTable = "PosStockOpname"
''            sField = "Docnum,DocDate,DocStatus,PosTranDate,Dscription,Refference,Note"
''            sfind = "Docnum"
'
'Case BrowseTables.BrowsStockOpname
'    AddToCombo "Docnum"
'    AddToCombo "DocDate"
'    AddToCombo "DocStatus"
'    AddToCombo "PosTranDate"
'    AddToCombo "Dscription"
'    AddToCombo "Reference"
'    AddToCombo "Note"
'    ClearGrid
'    AddColom "Doc.No", 1500
'    AddColom "Date", 1000
'    AddColom "Status", 1000
'    AddColom "Trans.Cut Off", 1500
'    AddColom "Dscription", 1500
'    AddColom "Refference", 3000
'    AddColom "Note", 3000
'    sTable = "PosStockOpname"
'
'Case BrowseTables.BrowsMultipleItem
'    AddToCombo "MultiItemID"
'    AddToCombo "Barcode"
'    AddToCombo "Active"
'    ClearGrid
'    AddColom "Multiple Item", 1500
'    AddColom "Barcode", 3000
'    AddColom "Active", 1500
'    sTable = "PosMultipleItem"
    
End Select
fCombo1(0).ListIndex = 0
fCombo1(1).ListIndex = 0

End Sub

Private Sub ClearGrid()
Dim i As Integer
If GridBrowse.Rows > 1 Then
    For i = GridBrowse.Rows - 1 To 1 Step -1
        GridBrowse.RemoveItem i
    Next
End If
End Sub

Private Sub AddToCombo(NamaField As String)
fCombo1(0).AddItem NamaField
If sFields <> "" Then
    sFields = sFields & "," & NamaField
Else
    sFields = NamaField
End If
End Sub

Private Sub AddColom(NamaField As String, lebar As Integer, posisi As PssText)
'fCombo1(0).AddItem NamaField
With GridBrowse
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = NamaField
    .COLWIDTH(.Cols - 1) = lebar
    '----
    '.ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(.Cols - 1) = posisi
End With
End Sub


Public Property Get YangDipilih() As Variant
    YangDipilih = svalue
End Property

Public Property Get Keterangan() As Variant
    Keterangan = sDesc
End Property


Private Sub Check1_Click()
MenuFrm.sCariotomatis = Check1.value
If Check1.value = 1 Then
    Command1(2).Enabled = False
Else
    Command1(2).Enabled = True
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    DoPilih
Case 1
    Unload Me
Case 2
    InitData sBrowse, sBrowsCriteria, sOrderBy, sDBaseConectionQ
End Select
End Sub

Private Sub fCombo1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
MsgBox "AA"
End Sub

Private Sub fCombo1_GotFocus(Index As Integer)
If Check1.value = 1 Then
    InitData sBrowse, sBrowsCriteria, sOrderBy, sDBaseConectionQ
End If
End Sub

Private Sub fCombo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MsgBox "fCombo1_KeyDown"
sListIndex = fCombo1(0).ListIndex
If KeyCode = 13 Then
    Sendkeys "{TAB}"
End If


End Sub


Private Sub DoPilih()
With GridBrowse
        svalue = .TextMatrix(.row, 0)
        If .Cols > 1 Then
        sDesc = .TextMatrix(.row, 1)
        End If
End With
    Me.Hide
End Sub

Private Sub fCombo1_Click(KeyCode As Integer)
If KeyCode = 13 Then
    Sendkeys "{TAB}"
End If

End Sub



Private Sub FlatButton1_Click(Index As Integer)

End Sub

Private Sub Form_Load()
Check1.value = MenuFrm.sCariotomatis
sListIndex = 1

End Sub

Private Sub GridBrowse_DblClick()
If GridBrowse.row > 0 Then
        DoPilih
    End If
End Sub

Private Sub GridBrowse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If GridBrowse.row > 0 Then
        DoPilih
    End If
End If
End Sub

Private Sub Text1_Change()
If Check1.value = 1 Then
    InitData sBrowse, sBrowsCriteria, sOrderBy, sDBaseConectionQ
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Sendkeys "{TAB}"
End If
End Sub



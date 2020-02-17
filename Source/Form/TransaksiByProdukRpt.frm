VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form TransaksiByProdukRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Sales By Product Report"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Penjualan"
      Height          =   1335
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   10455
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   19
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         FlatButton      =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   20
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         FlatButton      =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Filter Berdasarkan"
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Fungsi"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Kategori"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Brand"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Ditampilkan Untuk Transaksi"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Pengeluaran Lain-lain"
         Height          =   375
         Index           =   4
         Left            =   7320
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Pindah Antar Gudang"
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   22
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Penerimaan Lain-lain"
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Penjualan"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Pembelian"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Brand"
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   10455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tampil di Laporan"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Value           =   1  'Checked
         Width           =   3015
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
         Left            =   4080
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1080
         Width           =   6075
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
         Left            =   2220
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   1335
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
         Left            =   4080
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Width           =   6075
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "TransaksiByProdukRpt.frx":0000
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
         Index           =   1
         Left            =   3600
         TabIndex        =   14
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "TransaksiByProdukRpt.frx":001C
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
         Caption         =   "S/D Kode."
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Kode."
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Customer"
      Height          =   1935
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   10455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tampil di Laporan"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Value           =   1  'Checked
         Width           =   3015
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
         Index           =   7
         Left            =   2220
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   840
         Width           =   1335
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
         Index           =   6
         Left            =   4080
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   840
         Width           =   6075
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
         Index           =   5
         Left            =   2220
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1200
         Width           =   1335
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
         Left            =   4080
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1200
         Width           =   6075
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   29
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "TransaksiByProdukRpt.frx":0038
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
         Index           =   3
         Left            =   3600
         TabIndex        =   30
         Top             =   1200
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "TransaksiByProdukRpt.frx":0054
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
         Caption         =   "Dari Kode."
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "S/D Kode."
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   2055
      End
   End
End
Attribute VB_Name = "TransaksiByProdukRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim stransid, soption, scustoption As Integer
Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Dim sKriteria As String
Dim sstssiswa As String
If Option1(0).value = True Then
    sKriteria = "stssiswa='1'"
End If
If Option1(1).value = True Then
    sKriteria = "stssiswa='0'"
End If
If Option1(2).value = True Then
    sKriteria = "''=''"
End If

Select Case Index
Case 0
    Select Case soption
    Case 1
        oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    Case 2
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
    Case 3
        oBrowse.ShowFinder BrowsFunction, "", ubAscending, DBaseConection.Modul
    End Select
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    Select Case soption
    Case 1
        oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    Case 2
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
    Case 3
        oBrowse.ShowFinder BrowsFunction, "", ubAscending, DBaseConection.Modul
    End Select
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
Case 2
    Select Case scustoption
        Case 1
        oBrowse.ShowFinder BrowsSupplier, "", ubAscending, DBaseConection.Modul
    Case 2
        oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    Case 3
        oBrowse.ShowFinder BrowsSupplier, "", ubAscending, DBaseConection.Modul
    Case 4
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
    Case 5
        oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    End Select
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(6) = oBrowse.Keterangan
    End If
Case 3
    Select Case scustoption
    Case 1
        oBrowse.ShowFinder BrowsSupplier, "", ubAscending, DBaseConection.Modul
    Case 2
        oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    Case 3
        oBrowse.ShowFinder BrowsSupplier, "", ubAscending, DBaseConection.Modul
    Case 4
        oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
    Case 5
        oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    End Select
    If Not oBrowse.YangDipilih = "" Then
        Text1(5) = oBrowse.YangDipilih
        Text1(4) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Sales By Product Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransaksiByProdukRpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatCheckList 1, Me
oFormatOption 2, Me
FlatDatePicker1(0).value = DateSerial(Year(Now()), Month(Now()), 1)
FlatDatePicker1(1).value = Now()

istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width
soption = 1
scustoption = 2
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
Dim sstssiwa As String
Dim ssortby As Integer
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim stxtheaderQ As String
If Option2(0).value = True Then
    stransid = 1
    stxtheaderQ = "Transaksi Pembelian"
End If
If Option2(1).value = True Then
    stransid = 2
    stxtheaderQ = "Transaksi Penjualan"
End If
If Option2(2).value = True Then
    stransid = 3
    stxtheaderQ = "Transaksi Penerimaan Lain-lain"
End If
If Option2(3).value = True Then
    stransid = 5
    stxtheaderQ = "Transaksi Transfer Antar Gudang"
End If
If Option2(4).value = True Then
    stransid = 4
    stxtheaderQ = "Transaksi Keluar Lain-lain"
End If

stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")
ssortby = IIf(Option2(0).value = True, 1, 2)
If Text1(0).text = "" Then
    snoidsiswafr = oFindByQuery("select nodokumen from transaksi_keluar order by nodokumen asc limit 1 ", DBaseConection.Modul)
Else
    snoidsiswafr = Text1(0).text
End If
If Text1(2).text = "" Then
    snoidsiswato = oFindByQuery("select nodokumen from transaksi_keluar order by nodokumen desc limit 1 ", DBaseConection.Modul)
Else
    snoidsiswato = Text1(2).text
End If

sQuery = "call sp_transaksi_by_produk_rpt('"
sQuery = sQuery & stransid & "','"
sQuery = sQuery & stglfr & "','"
sQuery = sQuery & stglto & "','"
sQuery = sQuery & soption & "','"
sQuery = sQuery & Text1(0) & "','"
sQuery = sQuery & Text1(2) & "','"
sQuery = sQuery & scustoption & "','"
sQuery = sQuery & Text1(7) & "','"
sQuery = sQuery & Text1(5) & "',"

Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If
If stransid = 2 Or stransid = 1 Then
With arTransaksiByProdukRpt
    .lblHeaderTrx = stxtheaderQ

    If Check1(0).value = Checked Then
        .GroupHeader3.Visible = True
        .GroupFooter3.Visible = True
    Else
        .GroupHeader3.Visible = False
        .GroupFooter3.Visible = False
    End If
    If Check1(1).value = Checked Then
        .GroupHeader2.Visible = True
        .GroupFooter2.Visible = True
    Else
        .GroupHeader2.Visible = False
        .GroupFooter2.Visible = False
    End If
    
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery & "0)"
    If stransid = 1 Then
        .lblfee.Caption = ""
        .txtfee.DataField = ""
        .txtfee.text = ""
        '.txtfee1.DataField = ""
        '.txtfee1.Text = ""
        .txtfee2.text = ""
        .txtfee2.DataField = ""
        .txtfee3.text = ""
        .txtfee3.DataField = ""

    End If
'    .lblkode.Caption = "Kode Custmr"
'    .lblketerangan.Caption = "Nama Customer"
'    .txtkodeproduk.DataField = "custmrcode"
'    .txtproductname.DataField = "custmrname"
    .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
'    .PageSettings.LeftMargin = MenuFrm.skiri
'    .PageSettings.RightMargin = MenuFrm.skanan
    .Show
End With
End If

If Not (stransid = 2 Or stransid = 1) Then
With arTransaksiByProdukRpt2
    .lblHeaderTrx = stxtheaderQ

If Check1(0).value = Checked Then
        .GroupHeader3.Visible = True
        .GroupFooter3.Visible = True
    Else
        .GroupHeader3.Visible = False
        .GroupFooter3.Visible = False
    End If
    If Check1(1).value = Checked Then
        .GroupHeader2.Visible = True
        .GroupFooter2.Visible = True
    Else
        .GroupHeader2.Visible = False
        .GroupFooter2.Visible = False
    End If
    
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery & "0)"
    
'    .lblkode.Caption = "Kode Custmr"
'    .lblketerangan.Caption = "Nama Customer"
'    .txtkodeproduk.DataField = "custmrcode"
'    .txtproductname.DataField = "custmrname"
    .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
'    .PageSettings.LeftMargin = MenuFrm.skiri
'    .PageSettings.RightMargin = MenuFrm.skanan
    .Show
End With
End If

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Private Sub Option1_Click(Index As Integer)
Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Select Case Index
Case 0
    soption = 1
Case 1
    soption = 2
Case 2
    soption = 3
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
Text1(4) = ""
Text1(5) = ""
Text1(6) = ""
Text1(7) = ""
Select Case Index
Case 0
    scustoption = 1
    Frame1(4).Caption = "Supplier"
Case 1
    scustoption = 2
    Frame1(4).Caption = "Customer"
Case 2
    scustoption = 3
    Frame1(4).Caption = "Supplier"
Case 3
    scustoption = 4
    Frame1(4).Caption = "Customer"
    
Case 4
    scustoption = 5
    Frame1(4).Caption = ""
    Frame1(4).Enabled = False
    
End Select
End Sub

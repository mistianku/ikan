VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form LabaRugiByCustomerByProduct 
   BackColor       =   &H8000000A&
   Caption         =   "Laba Rugi Customer By Product Form"
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
      Caption         =   "Produk"
      Height          =   1335
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   3000
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
         Index           =   7
         Left            =   2220
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   360
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
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   360
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
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   720
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
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   720
         Width           =   6075
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   25
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LabaRugiByCustomerByProduct.frx":0000
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
         TabIndex        =   26
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LabaRugiByCustomerByProduct.frx":001C
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
         Caption         =   "Dari Kode.Produk"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "S/D Kode.Produk"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Penjualan"
      Height          =   1335
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   10455
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   17
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Ditampilkan "
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   10455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Termasuk Fee Nol"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   29
         Top             =   360
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Rekap"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Rinci"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Ditampilkan Rekap By"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Produk"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Customer"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Customer"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
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
         Index           =   3
         Left            =   4080
         TabIndex        =   14
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
         Index           =   2
         Left            =   2220
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   720
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
         Top             =   360
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
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LabaRugiByCustomerByProduct.frx":0038
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
         TabIndex        =   13
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LabaRugiByCustomerByProduct.frx":0054
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
         Caption         =   "S/D Kode.Customer"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Kode.Customer"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "LabaRugiByCustomerByProduct"
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
Dim sKriteria As String
Dim sstssiswa As String
If Option1(0).value = True Then
    sKriteria = "stssiswa='1'"
End If
If Option1(1).value = True Then
    sKriteria = "stssiswa='0'"
End If
'If Option1(2).value = True Then
'    sKriteria = "''=''"
'End If



Select Case Index
Case 0
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
Case 2
    oBrowse.ShowFinder BrowsMasterProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(6) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsMasterProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(5) = oBrowse.YangDipilih
        Text1(4) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Laba Rugi Customer By Product Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLabaRugiByCustomerByProductRpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 2, Me
oFormatCheckList 1, Me
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
Dim sfeenol As String
Dim ssortby As Integer
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim slblHeaderTrx As String

stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")
ssortby = IIf(Option2(0).value = True, 1, 2)
sstssiwa = IIf(Option1(0).value = True, "Y", "N")
'sfeenol = IIf(Check1(0).value = Unchecked, "N", "Y")
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

If ssortby = 1 Then
    slblHeaderTrx = "Report Fee By Customer"
    arJualRptbyPoduk.Label25.Caption = "Produk"
    If (Text1(0) = Text1(2)) And Not (Text1(0) = "" Or Text1(2) = "") Then
        arJualRptbyPoduk.GroupFooter1.Visible = False
    End If
Else
    slblHeaderTrx = "Report Fee By Produk"
    arJualRptbyPoduk.Label25.Caption = "Customer"
    If (Text1(7) = Text1(5)) And Not (Text1(7) = "" Or Text1(5) = "") Then
        arJualRptbyPoduk.GroupFooter1.Visible = False
    End If
End If
Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

If sstssiwa = "Y" Then
        sQuery = "call sp_transaksi_keluar_by_item_rpt_komisi('"
        sQuery = sQuery & stglfr & "','"
        sQuery = sQuery & stglto & "','"
        sQuery = sQuery & Text1(0) & "','"
        sQuery = sQuery & Text1(2) & "','"
        sQuery = sQuery & Text1(7) & "','"
        sQuery = sQuery & Text1(5) & "','"
        sQuery = sQuery & ssortby & "','"
        sQuery = sQuery & sfeenol & "',"
        
        'sp_transaksi_keluar_rpt`(IN stglfr DATE ,stglto DATE ,
        'skodecustomerfr CHAR(15),skodecustomerto CHAR(15),ssortby INT,sget INT)
        
        If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
            MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
            Exit Sub
        End If
        With arFeeRptbyPoduk
            .lblHeaderTrx = slblHeaderTrx
           
            .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
            .adoKu.Source = sQuery & "0)"
            .Label25.Caption = IIf(ssortby = "2", "Customer", "Produk")
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
Else
        sQuery = "call sp_transaksi_keluar_by_item_rpt_komisi_rkp('"
        sQuery = sQuery & stglfr & "','"
        sQuery = sQuery & stglto & "','"
        sQuery = sQuery & Text1(0) & "','"
        sQuery = sQuery & Text1(2) & "','"
        sQuery = sQuery & Text1(7) & "','"
        sQuery = sQuery & Text1(5) & "','"
        sQuery = sQuery & ssortby & "','"
        sQuery = sQuery & sfeenol & "',"
        
        'sp_transaksi_keluar_rpt`(IN stglfr DATE ,stglto DATE ,
        'skodecustomerfr CHAR(15),skodecustomerto CHAR(15),ssortby INT,sget INT)
        
        If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
            MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
            Exit Sub
        End If
        With arFeeRptbyPodukRekap
            .lblHeaderTrx = slblHeaderTrx
           
            .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
            .adoKu.Source = sQuery & "0)"
            .Label25.Caption = IIf(ssortby = "2", "Customer", "Produk")
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


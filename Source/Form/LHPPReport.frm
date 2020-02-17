VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form LHPPReport 
   BackColor       =   &H8000000A&
   Caption         =   "LHPP Report"
   ClientHeight    =   7905
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
   ScaleHeight     =   7905
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Dokumen"
      Height          =   975
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "All"
         Height          =   375
         Index           =   4
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Open"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Close"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   2775
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
      Caption         =   "Urut Berdasarkan"
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Kode Customer"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Tanggal"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Ditampilkan Secara"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Rinci"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Rekap"
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
      Caption         =   "kolektor"
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
         MouseIcon       =   "LHPPReport.frx":0000
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
         MouseIcon       =   "LHPPReport.frx":001C
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
         Caption         =   "S/D Kode.Kolektor"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Kode.Kolektor"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "LHPPReport"
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
    oBrowse.ShowFinder BrowsKolektor, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsKolektor, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub




Private Sub Form_Activate()
Dim sTitle As String
sTitle = "LHPP Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLHPPReport
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
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

Text1(0) = oFindByQuery("SELECT  MIN(kodekolektor) kodekolektor FROM master_kolektor", DBaseConection.Modul)
Text1(2) = oFindByQuery("SELECT  MAX(kodekolektor) kodekolektor FROM master_kolektor", DBaseConection.Modul)

Text1(1) = oFindByQuery("SELECT  namakolektor kodekolektor FROM master_kolektor where kodekolektor='" & Text1(0) & "'", DBaseConection.Modul)
Text1(3) = oFindByQuery("SELECT  namakolektor kodekolektor FROM master_kolektor where kodekolektor='" & Text1(2) & "'", DBaseConection.Modul)


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
Dim sdetail As String
Dim scustomercodefr As String
Dim scustomercodeto As String
Dim stglfr As String
Dim stglto As String
Dim sKriteria As String
Dim stxtkode As String
Dim stxtketerangan As String


Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
'CALL sp_transaksi_lhpp_print_insert_get('',15,'Admin');

'sQuery = "call sp_transaksi_lhpp_get_form('"
'sQuery = sQuery & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "','"
'sQuery = sQuery & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "','"
'sQuery = sQuery & Text1(0) & "','"
'sQuery = sQuery & Text1(0) & "',"

sQuery = "SELECT count(*) "
sQuery = sQuery & " FROM transaksi_lhpp a INNER JOIN transaksi_lhpp_detail1 a1 ON a.batchid=a1.batchid"
sQuery = sQuery & " WHERE a.tgldokumen between '" & Format(FlatDatePicker1(0).value, "YYYY-MM-DD") & "' and '" & Format(FlatDatePicker1(1).value, "YYYY-MM-DD") & "'"
sQuery = sQuery & " and a.kodekolektor between '" & Text1(0) & "' and '" & Text1(2) & "'"



If oFindByQuery(sQuery, DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If

sQuery = "SELECT upper(a.kodekolektor) as kodekolektor,a.nodokumen,a1.kodecustomer,a1.totnilkwitansi"
sQuery = sQuery & " FROM transaksi_lhpp a INNER JOIN transaksi_lhpp_detail1 a1 ON a.batchid=a1.batchid"
sQuery = sQuery & " WHERE a.tgldokumen between '" & Format(FlatDatePicker1(0).value, "YYYY-MM-DD") & "' and '" & Format(FlatDatePicker1(1).value, "YYYY-MM-DD") & "'"
sQuery = sQuery & " and a.kodekolektor between '" & Text1(0) & "' and '" & Text1(2) & "'"


With arlhppForm
   ' .lblHeaderTrx = "Form LHPP"
'    .lblCompany1 = MenuFrm.txtHeader(0)
'    .lblCompany2 = MenuFrm.txtHeader(1)
'    .lblCompany3 = MenuFrm.txtHeader(2)
    .Field5.text = "LHPP"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery ' & "0)"
    
'    .lblkode.Caption = "Kode Custmr"
'    .lblketerangan.Caption = "Nama Customer"
'    .txtkodeproduk.DataField = "custmrcode"
'    .txtproductname.DataField = "custmrname"

'    .PageSettings.Orientation = ddOPortrait
    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
'    .PageSettings.LeftMargin = MenuFrm.skiri
'    .PageSettings.RightMargin = MenuFrm.skanan
    .Show
    '.lblterbilang = terbilang.terbilang(100)
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


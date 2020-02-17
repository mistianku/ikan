VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MonitoringLHPPEntryRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Monitoring Penerimaan Pembayaran (LHPP) Report"
   ClientHeight    =   7320
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   11925
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
   ScaleHeight     =   7320
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kolektor"
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   120
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
         Index           =   4
         Left            =   2220
         TabIndex        =   22
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
         Index           =   5
         Left            =   4080
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   360
         Width           =   6075
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MonitoringLHPPEntryRpt.frx":0000
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
         Caption         =   "Kode.Kolektor"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   360
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
      Top             =   1080
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
      Caption         =   "Sort By"
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Customer"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   9
         Top             =   360
         Width           =   2175
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
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Ditampilkan "
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Rekap"
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Detail"
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
      Top             =   2520
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
         MouseIcon       =   "MonitoringLHPPEntryRpt.frx":001C
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
         MouseIcon       =   "MonitoringLHPPEntryRpt.frx":0038
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
Attribute VB_Name = "MonitoringLHPPEntryRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim txtFrame1, txtframe_fr, txtframe_to As String
Dim ssortby, sdetail As Integer
Dim lblLabelCutOff As String

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
         oBrowse.ShowFinder BrowsKolektor, "", ubAscending, DBaseConection.Modul

    If Not oBrowse.YangDipilih = "" Then
        Text1(4) = oBrowse.YangDipilih
        Text1(5) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub



Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Monitoring Penerimaan Pembayaran (LHPP) Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLHPPEntryRpt
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

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width

ssortby = 1
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
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim sstatusdok As String

stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")

sstatusdok = IIf(Option1(0).value = True, "N", IIf(Option1(1).value = True, "Y", "A"))
sdetail = IIf(Option2(0).value = True, 1, 2)
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

Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

If Text1(0) = "" And Text1(2) = "" Then
    lblLabelCutOff = " Semua Customer " & " Dari Tgl " & stglfr & " s.d " & stglto
Else
    lblLabelCutOff = " Dari Customer " & Text1(0) & " s.d " & Text1(1) & " Dari Tgl " & stglfr & " s.d " & stglto
End If
ssortby = IIf(Option1(0).value = True, 1, 2)

If sdetail = 1 Then
        sQuery = "call sp_monitoring_lhpp_entry('"
        sQuery = sQuery & Text1(4) & "','"
        sQuery = sQuery & stglfr & "','"
        sQuery = sQuery & stglto & "','"
        sQuery = sQuery & Text1(0) & "','"
        sQuery = sQuery & Text1(2) & "','"
        sQuery = sQuery & ssortby & "','"
        sQuery = sQuery & sstatusdok & "',"
        
        'sp_transaksi_keluar_rpt`(IN stglfr DATE ,stglto DATE ,
        'skodecustomerfr CHAR(15),skodecustomerto CHAR(15),ssortby INT,sget INT)
        
        If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
            MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
            Exit Sub
        End If
        With arLhppEntryDetail
            .lblHeaderTrx = "Penerimaan Penagihan (LHPP) Detail"
            .lblCutOffDate = lblLabelCutOff
            .Label50.Caption = IIf(ssortby = 1, "Tanggal", "Customer")
            
            .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
            .adoKu.Source = sQuery & "0)"
            .Label57 = IIf(ssortby = 1, "Sub Total Per Tanggal : ", "Sub Total Per Customer : ")
        
            '.PageSettings.Orientation = ddOPortrait
        
            .Show
        End With
Else
        sQuery = "call sp_monitoring_lhpp_entry_rekap('"
        sQuery = sQuery & Text1(4) & "','"
        sQuery = sQuery & stglfr & "','"
        sQuery = sQuery & stglto & "','"
        sQuery = sQuery & Text1(0) & "','"
        sQuery = sQuery & Text1(2) & "','"
        sQuery = sQuery & ssortby & "','"
        sQuery = sQuery & sstatusdok & "',"
        
        'sp_transaksi_keluar_rpt`(IN stglfr DATE ,stglto DATE ,
        'skodecustomerfr CHAR(15),skodecustomerto CHAR(15),ssortby INT,sget INT)
        
        If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
            MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
            Exit Sub
        End If
        With arLhppEntryRekap
            .lblHeaderTrx = "Penerimaan Penagihan (LHPP) Summary"
            .lblCutOffDate = lblLabelCutOff
            .Label50.Caption = IIf(ssortby = 1, "Tanggal", "Customer")
            
            .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
            .adoKu.Source = sQuery & "0)"
            .Label57 = IIf(ssortby = 1, "Sub Total Per Tanggal : ", "Sub Total Per Customer : ")
        
            '.PageSettings.Orientation = ddOPortrait
        
            .Show
        End With
End If
Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

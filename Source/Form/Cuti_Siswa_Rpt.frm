VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form Cuti_Siswa_Rpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
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
      Caption         =   "Status Dokumen"
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   10455
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Status Dokumen"
         Height          =   975
         Index           =   5
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   10455
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "Semua Status"
            Height          =   375
            Index           =   2
            Left            =   4080
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "Tidak Aktif"
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   23
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "Aktif"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Konfirmasi"
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Semua Status"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Sudah Konfirmasi"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Belum Konfirmasi"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2280
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Sort By Status Konfirmasi"
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Tidak"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Ya"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Master Siswa"
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   3600
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         MouseIcon       =   "Cuti_Siswa_Rpt.frx":0000
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
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Cuti_Siswa_Rpt.frx":001C
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
         Caption         =   "S/D No.ID.Siswa"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari No.ID.Siswa"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Cuti"
      Height          =   1215
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   10455
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   25
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
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
         TabIndex        =   26
         Top             =   720
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "S/D Tanggal"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Tanggal"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Cuti_Siswa_Rpt"
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
If Option1(2).value = True Then
    sKriteria = "''=''"
End If

Select Case Index
Case 0
    oBrowse.ShowFinder BrowsSiswa, sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsSiswa, sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Form Cuti Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCutiRpt
End Sub

Private Sub Form_Load()

oFormatOption 4, Me
istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

FlatDatePicker1(0).value = DateSerial(Year(Now), Month(Now), 1)
FlatDatePicker1(1).value = Now()
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
    Text1(i).Text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sstssiwa As String
Dim sdetail As String
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim ssortbykonfirmasists As String
Dim ssortbykonfirmsts As String
Dim sstatusdok As String

If Option4(0).value = True Then
    sstatusdok = "1"
End If
If Option4(1).value = True Then
    sstatusdok = "0"
End If
If Option4(2).value = True Then
    sstatusdok = "2"
End If

If Text1(0).Text = "" Then
    snoidsiswafr = oFindByQuery("select noidsiswa from master_siswa order by noidsiswa asc limit 1 ", parkir)
Else
    snoidsiswafr = Text1(0).Text
End If
If Text1(2).Text = "" Then
    snoidsiswato = oFindByQuery("select noidsiswa from master_siswa order by noidsiswa desc limit 1 ", parkir)
Else
    snoidsiswato = Text1(2).Text
End If

Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\cuti_rpt.Rpt"

Dim sKriteria As String
If Option1(0).value = True Then
sKriteria = " where konfirmasists  = '0'"
sstssiwa = 1
End If
If Option1(1).value = True Then
sKriteria = " where konfirmasists  = '1'"
sstssiwa = 0
End If
If Option1(2).value = True Then
sKriteria = " where '1'  = '1'"
sstssiwa = 2
End If

sKriteria = sKriteria & " and noidsiswa between '" & snoidsiswafr & "' and '" & snoidsiswato & "'"
If Option2(0).value = True Then
    ssortbykonfirmasists = "1"
Else
    ssortbykonfirmasists = "0"
End If
stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")
sKriteria = sKriteria & " and tanggal between '" & stglfr & "' and '" & stglto & "'"
sKriteria = sKriteria & " and dokumensts=" & IIf(sstatusdok = "2", "dokumensts", "'" & sstatusdok & "'")

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vtransaksi_cuti_rpt vtransaksi_cuti_rpt1 " & sKriteria & " "
'
'
Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(1) = "sortbykonfirmasists" & ";" & ssortbykonfirmasists & ";" & True
Me.cr1.ParameterFields(2) = "tglfr" & ";" & stglfr & ";" & True
Me.cr1.ParameterFields(3) = "tglto" & ";" & stglto & ";" & True
Me.cr1.ParameterFields(4) = "sortbykonfirmsts" & ";" & sstssiwa & ";" & True
Me.cr1.ParameterFields(5) = "noidsiswafr" & ";" & snoidsiswafr & ";" & True
Me.cr1.ParameterFields(6) = "noidsiswato" & ";" & snoidsiswato & ";" & True
Me.cr1.ParameterFields(7) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(8) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(9) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Form Cuti Siswa"
End Sub


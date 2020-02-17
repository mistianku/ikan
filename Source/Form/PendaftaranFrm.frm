VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form PendaftaranFrm 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   5835
   ClientLeft      =   5700
   ClientTop       =   3975
   ClientWidth     =   10170
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   6
      Left            =   240
      TabIndex        =   47
      Top             =   7320
      Width           =   9135
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
         Index           =   8
         Left            =   2235
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   240
         Width           =   6250
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
         Index           =   9
         Left            =   2235
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   600
         Width           =   6250
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Keterangan"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Referensi"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   2055
      End
   End
   Begin NeoCalendarII.DatePicker FlatDatePicker1 
      Height          =   315
      Index           =   0
      Left            =   6075
      TabIndex        =   44
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
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
   Begin Crystal.CrystalReport cr1 
      Left            =   9720
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSDFLATS.FlatButton BrowseUserID 
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   43
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      MouseIcon       =   "PendaftaranFrm.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Status Dokumen"
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Width           =   2175
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Manual"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   32
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Informasi Pendaftaran"
      ForeColor       =   &H80000008&
      Height          =   2175
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   5040
      Width           =   9135
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Penggolongan"
         Height          =   1455
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   4095
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Baru"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pindah Masuk"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Pelajaran Tambahan"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   960
            Width           =   3855
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Pelajaran"
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   3
         Left            =   4320
         TabIndex        =   12
         Top             =   600
         Width           =   4095
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "EFL"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   2280
            TabIndex        =   31
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "EE"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Bahasa Inggris"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   3855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H80000005&
            Caption         =   "Matematika"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   3855
         End
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   2
         Left            =   2235
         TabIndex        =   46
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
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
         Appearance      =   0  'Flat
         Caption         =   "Mulai Belajar "
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Informasi Siswa"
      ForeColor       =   &H0000FFFF&
      Height          =   3975
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   9135
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Tingkatan Sekolah "
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   8415
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "Diatas SMA"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   42
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "SMA"
            Height          =   255
            Index           =   3
            Left            =   2640
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "SMP"
            Height          =   255
            Index           =   2
            Left            =   1800
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "SD"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   39
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000A&
            Caption         =   "TK"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Manual"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1080
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Caption         =   "Perempuan"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Caption         =   "Laki Laki"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
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
         Left            =   2240
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3480
         Width           =   6250
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
         Left            =   2240
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3120
         Width           =   6250
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
         Left            =   2240
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2760
         Width           =   6250
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
         Left            =   2240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2400
         Width           =   6250
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
         Left            =   2240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   3555
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
         Left            =   2240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   6250
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
         Left            =   2240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   6250
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   1
         Left            =   5880
         TabIndex        =   45
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
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
         Appearance      =   0  'Flat
         Caption         =   "Telepon"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Alamat Lengkap"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Kelas"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Sekolah"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Tempat.Tgl Lahir"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Jenis Kelamin"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Nama Lengkap"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "No ID"
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
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
      Left            =   6075
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "No Pendaftaran"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   13
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal "
      Height          =   315
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   2565
   End
End
Attribute VB_Name = "PendaftaranFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim skode As String
Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm

Dim sInsertDetail As String
Dim sUpdateDetail As String

Dim snopendaftaran As String
Dim stglpendaftaran As String
Dim sketerangan As String
Dim sreferensi As String
Dim snoidsiswa As String
Dim spenggolongan As String
Dim spelajaran As String
Dim saudituser As String
Dim sauditdate As Date
Dim stglmasuk As String
Dim snmlengkap As String
Dim sjnskelamin As String
Dim stptlahir As String
Dim stgllahir As String
Dim sKelas As String
Dim saslsekolah As String
Dim salmtrumah1 As String
Dim snotelprumah As String
Dim stingkatansklh As String

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from vtransaksi_pendaftaran where nopendaftaran='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select *  from vtransaksi_pendaftaran order by nopendaftaran asc limit 1"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub

Public Sub MoveNext()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select  *  from vtransaksi_pendaftaran where nopendaftaran >'" & Text1(0).Text & "' order by nopendaftaran asc limit 1"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveNext"
End Sub
Public Sub MovePrevious()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select  *  from vtransaksi_pendaftaran where nopendaftaran<'" & Text1(0).Text & "' order by nopendaftaran desc limit 1"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MovePrevious"
End Sub

Public Sub MoveLast()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select *  from vtransaksi_pendaftaran order by nopendaftaran desc limit 1 "
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    Else
        cleardata
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveLast"
End Sub
Public Sub SaveData()
Dim ires As Integer
If Check2(0).value = 1 And Text1(0).Text = "" Then
   MsgBox "Isi No.Pendaftaran Siswa Secara Manual ! ", vbInformation
   Exit Sub
End If

If Check2(1).value = 1 And Text1(1).Text = "" Then
   MsgBox "Isi No.ID Siswa Secara Manual ! ", vbInformation
   Exit Sub
End If

    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran
        End If
    End If
End Sub
Public Sub DeleteData()
    Dim ires As Integer
    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
    If ires = 6 Then
        If DoDeleteData Then
             MsgBox "Data Sudah Terhapus", , "Hapus Data"
             MovePrevious
        End If
    End If
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
         oCon.BeginTrans

        If istatus = StatusForm.DataBaru Then
            oCon.Execute sInsertDetail
            sQuery = sInsert
        Else
            oCon.Execute sUpdateDetail
            sQuery = sUpdate
        End If
        oCon.Execute sQuery
        oCon.CommitTrans
        oCon.Close
        DoSaveData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
oCon.RollbackTrans
oCon.Close
MainModule.ShowMessage Err.Description, "savedata"
End Function
Private Function DoDeleteData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "Delete from transaksi_pendaftaran where nopendaftaran='" & snopendaftaran & "'"
        oCon.Execute sQuery
        oCon.Close
        DoDeleteData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    If MenuFrm.sAplikasiDemo Then
        If oCekJumlahTrx("transaksi_pendaftaran", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran
    Text1(0).Locked = False
    
    If oFindByQuery("select autonodefault from setting_document where docid=" & transaksi_pendaftaran, parkir) = "0" Then
    Check2(0).value = 1
    Else
    Check2(0).value = 0
    End If
    
    If Check2(0).value = 1 Then
        Text1(0).SetFocus
    Else
        Text1(2).SetFocus
    End If
    
    Check3(0).value = 1
    Check3(1).value = 0
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    FlatDatePicker1(0).value = Now()
    FlatDatePicker1(1).value = Now()
    FlatDatePicker1(2).value = Now()
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
If istatus = StatusForm.DataBaru Then
    snopendaftaran = ToText(IIf(Text1(0).Text = "", GetDocnum(transaksi_pendaftaran, True, parkir), Text1(0).Text))
    Text1(0).Text = snopendaftaran
    snoidsiswa = ToText(IIf(Trim(Text1(1).Text) = "", GetDocnum(master_siswa, True, parkir), Trim(ToText(Text1(1).Text))))
    Text1(1).Text = snoidsiswa
Else
    snopendaftaran = Text1(0).Text
    snoidsiswa = Text1(1).Text
End If

stglpendaftaran = Format(FlatDatePicker1(0).value, "yyyy/mm/dd")
sketerangan = Replace(Text1(8).Text, "'", "\'")
sreferensi = Replace(Text1(9).Text, "'", "\'")
If Option4(0).value = True Then
    stingkatansklh = "0"
End If
If Option4(1).value = True Then
    stingkatansklh = "1"
End If
If Option4(2).value = True Then
    stingkatansklh = "2"
End If
If Option4(3).value = True Then
    stingkatansklh = "3"
End If
If Option4(4).value = True Then
    stingkatansklh = "4"
End If

stglmasuk = Format(FlatDatePicker1(2).value, "yyyy/mm/dd")
snmlengkap = Replace(ToText(Text1(2).Text), "'", "\'")
sjnskelamin = IIf(Option3(0).value = True, "L", "P")
stptlahir = ToText(Text1(3).Text)
stgllahir = Format(FlatDatePicker1(1).value, "yyyy/mm/dd")
sKelas = ToText(Text1(5).Text)
saslsekolah = ToText(Text1(4).Text)
salmtrumah1 = ToText(Text1(6).Text)
snotelprumah = ToText(Text1(7).Text)
spenggolongan = IIf(Option1(0).value = True, "1", IIf(Option1(1).value = True, "2", "3"))
spelajaran = IIf(Option2(0).value = True, "1", IIf(Check1(0).value = 1, "2", "3"))
saudituser = MenuFrm.sUserID
sauditdate = Format(Now(), "yyyy/mm/dd")


    sUpdate = "update transaksi_pendaftaran "
    sUpdate = sUpdate & "set "
    sUpdate = sUpdate & "nopendaftaran='" & snopendaftaran & "',"
    sUpdate = sUpdate & "tglpendaftaran='" & stglpendaftaran & "',"
    sUpdate = sUpdate & "keterangan='" & sketerangan & "',"
    sUpdate = sUpdate & "referensi='" & sreferensi & "',"
    sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "',"
    sUpdate = sUpdate & "penggolongan='" & spenggolongan & "',"
    sUpdate = sUpdate & "pelajaran='" & spelajaran & "',"
    sUpdate = sUpdate & "tingkatansklh='" & stingkatansklh & "',"
    sUpdate = sUpdate & "audituser='" & saudituser & "',"
    sUpdate = sUpdate & "auditdate='" & sauditdate & "'"
    sUpdate = sUpdate & "where "
    sUpdate = sUpdate & "nopendaftaran='" & snopendaftaran & "'"
    
    sInsert = "insert into transaksi_pendaftaran"
    sInsert = sInsert & "("
    sInsert = sInsert & "nopendaftaran,"
    sInsert = sInsert & "tglpendaftaran,"
    sInsert = sInsert & "keterangan,"
    sInsert = sInsert & "referensi,"
    sInsert = sInsert & "noidsiswa,"
    sInsert = sInsert & "penggolongan,"
    sInsert = sInsert & "pelajaran,tingkatansklh,"
    sInsert = sInsert & "audituser,"
    sInsert = sInsert & "auditdate"
    sInsert = sInsert & ")"
    sInsert = sInsert & " values ('"
    sInsert = sInsert & snopendaftaran & "','"
    sInsert = sInsert & stglpendaftaran & "','"
    sInsert = sInsert & sketerangan & "','"
    sInsert = sInsert & sreferensi & "','"
    sInsert = sInsert & snoidsiswa & "','"
    sInsert = sInsert & spenggolongan & "','"
    sInsert = sInsert & spelajaran & "','"
    sInsert = sInsert & stingkatansklh & "','"
    sInsert = sInsert & saudituser & "','"
    sInsert = sInsert & sauditdate & "')"
    
    oSet_master_siswa snoidsiswa
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsPendaftaran, ""
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
End Select
Set oBrowse = Nothing
End Sub

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    If Check1(0).value = 1 Then
        Check1(1).value = 0
    Else
        Check1(1).value = 1
    End If
    
Case 1
    If Check1(1).value = 1 Then
        Check1(0).value = 0
    Else
        Check1(0).value = 1
    End If
End Select

End Sub

Private Sub Check2_Click(Index As Integer)

    Text1(0).Enabled = IIf(Check2(0).value = 1, True, False)
    Text1(1).Enabled = IIf(Check2(1).value = 1, True, False)
    Select Case Index
    Case 0
        Select Case Check2(0).value
        Case 1
            Text1(0).SetFocus
        Case 0
            Text1(2).SetFocus
        End Select
    Case 1
        
        Select Case Check2(1).value
        Case 1
            Text1(1).SetFocus
        Case 0
            Text1(2).SetFocus
        End Select
    End Select

End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Form Pendaftaran"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPendaftaran

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
oFormatOption 4, Me
oFormatCheckList 3, Me
cleardata
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    
    If oRs("dokstatus") = "O" Then
        Check3(0).value = 1
        Check3(1).value = 0
    Else
        Check3(0).value = 0
        Check3(1).value = 1
    End If
    
    Text1(0).Enabled = IIf(Check2(0).value = 1, True, False)
    Text1(0).Text = ToText(oRs("nopendaftaran"))
    KodeUserAksesTemp = oRs("nopendaftaran")
    Text1(0).Locked = True
    
    FlatDatePicker1(0).value = oRs("tglpendaftaran")
    Text1(1).Enabled = IIf(Check2(1).value = 1, True, False)
    Text1(1).Text = ToText(oRs("noidsiswa"))
    Text1(2).Text = ToText(oRs("nmlengkap"))
    Select Case oRs("jnskelamin")
    Case "L"
        Option3(0).value = True
    Case "P"
        Option3(1).value = True
    End Select
    
    Option4(ToNumber(oRs("tingkatansklh"))).value = True
        
    Text1(3).Text = ToText(oRs("tptlahir"))
     FlatDatePicker1(1).value = ToText(oRs("tgllahir"))
    Text1(4).Text = ToText(oRs("aslsekolah"))
    Text1(5).Text = ToText(oRs("kelas"))
    Text1(6).Text = ToText(oRs("almtrumah1"))
    Text1(7).Text = ToText(oRs("notelprumah"))
     FlatDatePicker1(2).value = oRs("tglmasuk")
    '=oRs("penggolongan")
    Select Case oRs("penggolongan")
    Case "1"
        Option1(0).value = True
    Case "2"
        Option1(1).value = True
    Case "3"
        Option1(2).value = True
    End Select
    '=oRs("pelajaran")
    Select Case oRs("pelajaran")
    Case "1"
        Option2(0).value = True
    Case "2"
        Option2(1).value = True
        Check1(0).value = 1
    Case "3"
        Option2(1).value = True
        Check1(1).value = 1
    End Select
    Text1(8).Text = oRs("keterangan")
    Text1(9).Text = oRs("referensi")

    
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub

Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
    Check1(0).Enabled = False
    Check1(1).Enabled = False
Case 1
    Check1(0).Enabled = True
    Check1(1).Enabled = True
    If Check1(0).value = 0 And Check1(1).value = 0 Then
        Check1(0).value = 1
    End If
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000B
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).Text
End Sub


Public Function oSet_master_siswa(snoid_siswa As String)
Dim iNoidSiswa As String

sInsertDetail = "insert into master_siswa "
sInsertDetail = sInsertDetail & "("
sInsertDetail = sInsertDetail & "noidsiswa,"
sInsertDetail = sInsertDetail & "nmlengkap,"
sInsertDetail = sInsertDetail & "jnskelamin,"
sInsertDetail = sInsertDetail & "tptlahir,"
sInsertDetail = sInsertDetail & "tgllahir,"
sInsertDetail = sInsertDetail & "aslsekolah,"
sInsertDetail = sInsertDetail & "kelas,"
sInsertDetail = sInsertDetail & "almtrumah1,"
sInsertDetail = sInsertDetail & "notelprumah,"
sInsertDetail = sInsertDetail & "tglmasuk,tingkatansklh,"
sInsertDetail = sInsertDetail & "audituser,"
sInsertDetail = sInsertDetail & "auditdate"
sInsertDetail = sInsertDetail & ")"
sInsertDetail = sInsertDetail & " values "
sInsertDetail = sInsertDetail & "('"
sInsertDetail = sInsertDetail & snoid_siswa & "','"
sInsertDetail = sInsertDetail & snmlengkap & "','"
sInsertDetail = sInsertDetail & sjnskelamin & "','"
sInsertDetail = sInsertDetail & stptlahir & "','"
sInsertDetail = sInsertDetail & stgllahir & "','"
sInsertDetail = sInsertDetail & saslsekolah & "','"
sInsertDetail = sInsertDetail & sKelas & "','"
sInsertDetail = sInsertDetail & salmtrumah1 & "','"
sInsertDetail = sInsertDetail & snotelprumah & "','"
sInsertDetail = sInsertDetail & stglmasuk & "','"
sInsertDetail = sInsertDetail & stingkatansklh & "','"
sInsertDetail = sInsertDetail & saudituser & "','"
sInsertDetail = sInsertDetail & sauditdate & "'"
sInsertDetail = sInsertDetail & ")"


sUpdateDetail = "Update master_siswa set "
sUpdateDetail = sUpdateDetail & "nmlengkap='" & snmlengkap & "',"
sUpdateDetail = sUpdateDetail & "jnskelamin='" & sjnskelamin & "',"
sUpdateDetail = sUpdateDetail & "tptlahir='" & stptlahir & "',"
sUpdateDetail = sUpdateDetail & "tgllahir='" & stgllahir & "',"
sUpdateDetail = sUpdateDetail & "aslsekolah='" & saslsekolah & "',"
sUpdateDetail = sUpdateDetail & "kelas='" & sKelas & "',"
sUpdateDetail = sUpdateDetail & "almtrumah1='" & salmtrumah1 & "',"
sUpdateDetail = sUpdateDetail & "notelprumah='" & snotelprumah & "',"
sUpdateDetail = sUpdateDetail & "tglmasuk='" & stglmasuk & "',"
sUpdateDetail = sUpdateDetail & "tingkatansklh='" & stingkatansklh & "',"
sUpdateDetail = sUpdateDetail & "audituser='" & saudituser & "',"
sUpdateDetail = sUpdateDetail & "auditdate='" & sauditdate & "'"
sUpdateDetail = sUpdateDetail & "where "
sUpdateDetail = sUpdateDetail & "noidsiswa='" & snoid_siswa & "'"

End Function

Public Sub Execution()
On Error GoTo errhandler
Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\PendaftaranFrm.Rpt"

Dim sKriteria As String
sQuery = "select * from vtransaksi_pendaftaran_rpt vtransaksi_pendaftaran_rpt1 where nopendaftaran='" & Text1(0) & "'"

Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(1) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(2) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(3) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Form Pendaftaran"
End Sub

Private Sub oTextBorderStyleNone()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).BorderStyle = 0
    Text1(i).Refresh
Next
End Sub


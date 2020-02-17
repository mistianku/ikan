VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Preference_special_Frm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Prefernce Spesial Form"
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
   Begin MSComDlg.CommonDialog cd1 
      Left            =   960
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "txt"
      InitDir         =   "c:\*.txt"
   End
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   7560
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "Preference_special_Frm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Export Preference To Text"
   End
   Begin VB.Frame frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Setting Faktur"
      ForeColor       =   &H80000008&
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   11775
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
         Index           =   20
         Left            =   2340
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   2820
         Width           =   7155
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tampil Logo"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Entrian Format Indonesia"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   6360
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   3135
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   2340
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         Index           =   19
         Left            =   2340
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3435
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
         Index           =   18
         Left            =   2340
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   2760
         Visible         =   0   'False
         Width           =   3435
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
         Index           =   17
         Left            =   2340
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2400
         Visible         =   0   'False
         Width           =   7155
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
         Index           =   11
         Left            =   2340
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   600
         Width           =   7155
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
         Index           =   12
         Left            =   2340
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   960
         Width           =   7155
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
         Index           =   13
         Left            =   2340
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1320
         Width           =   7155
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
         Index           =   14
         Left            =   2340
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1680
         Width           =   7155
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
         Index           =   15
         Left            =   2340
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3435
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
         Index           =   16
         Left            =   5820
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   38
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Preference_special_Frm.frx":001C
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
         Caption         =   "Nama File Logo .JPG"
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   45
         Top             =   2820
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Model Cetak Invoice"
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cost Default"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Provinsi"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tgl.NPWP"
         Height          =   315
         Index           =   14
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No Rekening"
         Height          =   315
         Index           =   13
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Bank Cabang"
         Height          =   315
         Index           =   12
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Atas Nama"
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Usaha"
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   27
         Top             =   1680
         Width           =   2055
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Alamat Utama"
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   11775
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
         Index           =   10
         Left            =   5820
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2040
         Width           =   3675
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
         Left            =   2340
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2040
         Width           =   3435
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
         Index           =   8
         Left            =   5820
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1680
         Width           =   3675
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
         Left            =   2340
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1680
         Width           =   3435
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
         Left            =   2340
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1320
         Width           =   7155
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
         Left            =   5820
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   3675
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
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   3435
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
         Left            =   2340
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   7155
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
         Left            =   2340
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Width           =   7155
      End
      Begin VB.Label Label1 
         Caption         =   "Faximale/Email "
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Telp1,Telp2"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Provinsi"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kota / Kode Pos"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
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
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   7155
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
         Left            =   2340
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   42
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Preference_special_Frm.frx":0038
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
         Caption         =   "Nama Instansi"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Instansi"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Preference_special_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sid As Integer
Dim sCmpnyID As String
Dim sCmnyName As String
Dim sAddress1 As String
Dim sAddress2 As String
Dim sCity As String
Dim sZipCode As String
Dim sState As String
Dim sPhone1 As String
Dim sPhone2 As String
Dim sFaximale As String
Dim sEmailAddress As String
Dim sNPWP As String
Dim sNPWPDate As String
Dim sPKPName As String
Dim sPKPAddress1 As String
Dim sPKPAddress2 As String
Dim sPKPCity As String
Dim sPKPZipCode As String
Dim sPKPState As String
Dim sCostDefault As String
Dim sPrintInvMode As String
Dim sisIndonesianFormat As String

Dim sKodeUserAkses As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm
Dim sis_image As Integer
Dim simage_name As String
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_preferences_special_get('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_preferences_special_get('" & sKodeUserAkses & "',1)"
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
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_preferences_special_get('" & sKodeUserAkses & "',3)"
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
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_preferences_special_get('" & sKodeUserAkses & "',2)"
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
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_preferences_special_get('" & sKodeUserAkses & "',4)"
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
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.sisIndonesianFormat = sisIndonesianFormat
             
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = StatusForm.DataBaru Then
            sQuery = sInsert
            oCon.Execute sQuery
        Else
            sQuery = sUpdate
            oCon.Execute sQuery
        End If
        
        oCon.Execute "call sp_master_customer_company_insert_from_customer_all('" & MenuFrm.sUserID & "')"
        oCon.Close
        DoSaveData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "savedata"
End Function
Private Function DoDeleteData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = sDelete
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
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
sCmpnyID = (Text1(0))
sCmnyName = (Text1(1))
sAddress1 = (Text1(2))
sAddress2 = (Text1(3))
sCity = (Text1(4))
sZipCode = (Text1(5))
sState = (Text1(6))
sPhone1 = (Text1(7))
sPhone2 = (Text1(8))
sFaximale = (Text1(9))
sEmailAddress = ToText(Text1(10))
sNPWP = ToText(Text1(11))
sNPWPDate = Format(FlatDatePicker1.value, "yyyy-MM-DD")
sPKPName = ToText(Text1(12))
sPKPAddress1 = ToText(Text1(13))
sPKPAddress2 = ToText(Text1(14))
sPKPCity = ToText(Text1(15))
sPKPZipCode = ToText(Text1(16))
sPKPState = Text1(17)
sCostDefault = Text1(18)
sPrintInvMode = Text1(19)
simage_name = Text1(20)
If Check1(0) = 1 Then
sisIndonesianFormat = "Y"
Else
sisIndonesianFormat = "N"
End If

If Check1(1) = 1 Then
sis_image = "1"
Else
sis_image = "0"
End If

sInsert = "call sp_master_preferences_special_insert("
sQuery = "'" & sid & "',"
sQuery = sQuery & "'" & sCmpnyID & "',"
sQuery = sQuery & "'" & sCmnyName & "',"
sQuery = sQuery & "'" & sAddress1 & "',"
sQuery = sQuery & "'" & sAddress2 & "',"
sQuery = sQuery & "'" & sCity & "',"
sQuery = sQuery & "'" & sZipCode & "',"
sQuery = sQuery & "'" & sState & "',"
sQuery = sQuery & "'" & sPhone1 & "',"
sQuery = sQuery & "'" & sPhone2 & "',"
sQuery = sQuery & "'" & sFaximale & "',"
sQuery = sQuery & "'" & sEmailAddress & "',"
sQuery = sQuery & "'" & sNPWP & "',"
sQuery = sQuery & "'" & sNPWPDate & "',"
sQuery = sQuery & "'" & sPKPName & "',"
sQuery = sQuery & "'" & sPKPAddress1 & "',"
sQuery = sQuery & "'" & sPKPAddress2 & "',"
sQuery = sQuery & "'" & sPKPCity & "',"
sQuery = sQuery & "'" & sPKPZipCode & "',"
sQuery = sQuery & "'" & sPKPState & "',"
sQuery = sQuery & "'" & sCostDefault & "',"
sQuery = sQuery & "'" & sPrintInvMode & "',"
sQuery = sQuery & "'" & sisIndonesianFormat & "',"
sQuery = sQuery & "'" & MenuFrm.sUserID & "',"
sQuery = sQuery & "'" & sis_image & "',"
sQuery = sQuery & "'" & simage_name & "')"

sInsert = sInsert & sQuery
sUpdate = Replace(sInsert, "insert", "update")
sDelete = Replace(sInsert, "insert", "delete")
     
    
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 1
    oBrowse.ShowFinder BrowsMasterPreferencesSpecial, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 0
    oBrowse.ShowFinder BrowsHarga, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(18) = oBrowse.YangDipilih
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Prefernce Spesial"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPreference_special_Frm
BrowseUserID(0).Top = Text1(18).Top
BrowseUserID(0).Height = Text1(18).Height
BrowseUserID(0).Left = Text1(18).Left + Text1(18).Width
BrowseUserID(1).Top = Text1(0).Top
BrowseUserID(1).Height = Text1(0).Height
BrowseUserID(1).Left = Text1(0).Left + Text1(0).Width

End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatCheckList 1, Me
cd1.DialogTitle = "Pilih Lokasi File Txt Ditaruh !! "
cd1.FileName = App.Path & "\*.txt"
cleardata

istatus = Normal
MoveLast
'Text1(2).SetFocus
'    Text1(0).TabIndex = 0
'    Text1(1).TabIndex = 1
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    sid = (oRs("id"))
    Text1(0).text = (oRs("CmpnyID"))
    sKodeUserAkses = Text1(0).text
    KodeUserAksesTemp = (oRs("CmpnyID"))
    'Text1(0).Locked = True
    Text1(1).text = (oRs("CmnyName"))
    Text1(0) = (oRs("CmpnyID"))
    Text1(1) = (oRs("CmnyName"))
    Text1(2) = (oRs("Address1"))
    Text1(3) = (oRs("Address2"))
    Text1(4) = (oRs("City"))
    Text1(5) = (oRs("ZipCode"))
    Text1(6) = (oRs("State"))
    Text1(7) = (oRs("Phone1"))
    Text1(8) = (oRs("Phone2"))
    Text1(9) = (oRs("Faximale"))
    Text1(10) = oRs("EmailAddress")
    Text1(11) = oRs("NPWP")
    FlatDatePicker1.value = oRs("NPWPDate")
    Text1(12) = oRs("PKPName")
    Text1(13) = oRs("PKPAddress1")
    Text1(14) = oRs("PKPAddress2")
    Text1(15) = oRs("PKPCity")
    Text1(16) = oRs("PKPZipCode")
    Text1(17) = oRs("PKPState")
    Text1(18) = oRs("CostDefault")
    Text1(19) = oRs("PrintInvMode")
    Text1(20) = oRs("image_name")

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    
    If oRs("isIndonesianFormat") = "Y" Then
        Check1(0).value = 1
    Else
         Check1(0).value = 0
    End If
    
    If oRs("is_image") = "1" Then
        Check1(1).value = 1
    Else
         Check1(1).value = 0
    End If
    
    
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub

Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).text = ""
Next
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
'If Index = 0 Then FindData Text1(0).text
End Sub

Public Sub oBackupPrefernce(sNamaFile As String)
Dim sConKu As New ADODB.Connection
Dim sRstku As New ADODB.Recordset
Dim sQuery As String
'Dim sConnectku  As Integer

sQuery = "SELECT * INTO OUTFILE '" & sNamaFile & "'"
sQuery = sQuery & " FIELDS TERMINATED BY  ';' "
sQuery = sQuery & " LINES TERMINATED BY '\n' FROM master_preferences "


If sConKu.State = 1 Then sConKu.Close
sConKu.Open MainModule.Conectionku(DBaseConection.Modul)
sConKu.Execute (sQuery)

sConKu.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, " oBackupPrefernce "

End Sub
Private Sub FlatButton1_Click()
On Error GoTo errhandler
Dim sNamaFile As String
Dim sNamaFile2 As String
cd1.FileName = "*.txt"
cd1.ShowSave

If cd1.FileName = "*.txt" Then Exit Sub
sNamaFile = Replace(UCase(cd1.FileName), ".TXT", "", DBaseConection.Modul) & ".txt"
sNamaFile2 = Dir(sNamaFile, vbReadOnly)
If sNamaFile2 = "" Then
Else
Kill sNamaFile
End If
WriteData sNamaFile
MsgBox "Export File Txt Selesai ", vbInformation
Exit Sub
errhandler:

End Sub

Public Sub WriteData(sNamaFile As String)
Dim sConKu As New ADODB.Connection
Dim sRstku As New ADODB.Recordset
Dim sQuery As String
Dim s As New FileSystemObject
Dim s1 As TextStream
'sTxtTujuan = App.Path & "\" & InputBox("Ketik Nama File Export !!", "Cuman Ngingetin !!") & ".SQL"
s.CreateTextFile sNamaFile
s.OpenTextFile sNamaFile, ForWriting, True
Set s1 = s.OpenTextFile(sNamaFile, ForWriting, True)

If sConKu.State = 1 Then sConKu.Close
sConKu.Open MainModule.Conectionku(DBaseConection.Modul)
Dim sFields As Integer
Dim sTextHeader As String
Dim jBrs As Integer
Dim sPesanTxt As String
    sQuery = "Select count(*) from master_preferences "
    Set sRstku = sConKu.Execute(sQuery)

    sQuery = "select *  from master_preferences "
    Set sRstku = sConKu.Execute(sQuery)
    sFields = sRstku.Fields.Count
    
    Dim irow As Double
    irow = 1
    
    Dim sString As String
    Dim iAwal As Integer
    Dim sKeluar As Integer
    sKeluar = 0
    iAwal = 1
    Do While Not sRstku.EOF

    If irow = 1 Then
        'sTextHeader = " insert into master_preferences  select "
        sTextHeader = ""
        sString = sTextHeader
    End If
    Dim iField As Integer
    Dim sNumber As Double

'   sString = sString & "("
    For iField = 0 To sFields - 1
        'Debug.Print sRstku.Fields(iField).Name
        Select Case sRstku.Fields(iField).Type
        Case adNumeric, adInteger
            sNumber = IIf(IsNull(sRstku(iField)), 0, sRstku(iField))
            sString = sString & sNumber & ","
        Case adDBTimeStamp
            sString = sString & "'" & IIf(IsNull(sRstku(iField)), "2001-01-01", Format(sRstku(iField), "yyyy-mm-dd")) & "',"
        Case Else
            sString = sString & "'" & IIf(IsNull(sRstku(iField)), "", Replace(IIf(IsNull(sRstku(iField)), "", sRstku(iField)), "'", "", DBaseConection.Modul)) & "',"
        End Select

    Next
    sString = sString & "),"
   '-----
    
     
    sRstku.MoveNext
    If irow = jBrs Then
'        pBar.Refresh
        sKeluar = 1
        
        For irow = 1 To 1 '--- kasih jarak ---
            s1.WriteLine ""
        Next
        irow = 1
        sString = Replace(sString, ",),", ";")
        s1.WriteLine sString
        If Not sRstku.EOF Then
                sString = sTextHeader
                sKeluar = 0
        Else
            
        End If
    End If
    irow = irow + 1
    '---- check record teralhir ----
        If sRstku.EOF And sKeluar = 0 Then
            sString = Replace(sString, ",),", "", DBaseConection.Modul)
            s1.WriteLine sString
            
        Else
            sString = Replace(sString, ",),", "),")
        End If

    
    Loop
    
    
    For irow = 1 To 3 '--- kasih jarak ---
        s1.WriteLine ""
    Next

lanjut:

s1.Close
    
    
sConKu.Close
'End With

End Sub



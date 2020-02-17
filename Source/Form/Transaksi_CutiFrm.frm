VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form Transaksi_CutiFrm 
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
   Begin Crystal.CrystalReport cr1 
      Left            =   5160
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Konfirmasi"
      Height          =   1575
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   6960
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Masuk Kembali"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   40
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Status Siswa"
         Height          =   615
         Index           =   5
         Left            =   8640
         TabIndex        =   35
         Top             =   360
         Width           =   2655
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Cuti"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Aktif"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Konfirmasi"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   29
         Top             =   360
         Width           =   1575
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
         Left            =   2160
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   720
         Width           =   4815
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   4
         Left            =   2160
         TabIndex        =   34
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   5
         Left            =   2160
         TabIndex        =   39
         Top             =   1080
         Width           =   4815
         _ExtentX        =   8493
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
         Enabled         =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.Label Label1 
         Caption         =   "Msk. Kembali Tanggal"
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Terdaftar di Kelas"
      Height          =   2295
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   4560
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   11535
         _cx             =   20346
         _cy             =   3201
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   10485760
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Transaksi_CutiFrm.frx":0000
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Informasi Cuti"
      Height          =   1815
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2640
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
         Index           =   4
         Left            =   2160
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   960
         Width           =   1035
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   2
         Left            =   2160
         TabIndex        =   32
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   3
         Left            =   2160
         TabIndex        =   33
         Top             =   1320
         Width           =   4815
         _ExtentX        =   8493
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
         Caption         =   "Masuk Tanggal"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Hari"
         Height          =   315
         Index           =   8
         Left            =   3240
         TabIndex        =   19
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai Tanggal"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai Tanggal"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Mulai Tanggal"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status"
      Height          =   615
      Index           =   1
      Left            =   9120
      TabIndex        =   11
      Top             =   120
      Width           =   2775
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Informasi Dokumen"
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   9360
         TabIndex        =   30
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Manual"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   8280
         TabIndex        =   28
         Top             =   240
         Width           =   975
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
         Height          =   765
         Index           =   3
         Left            =   2160
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   960
         Width           =   4875
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
         Left            =   2160
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   600
         Width           =   4875
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
         Left            =   2160
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   4275
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   9360
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1870
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   11280
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Transaksi_CutiFrm.frx":00D2
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
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Transaksi_CutiFrm.frx":00EE
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
         Caption         =   "Alamat"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "N a m a"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "No.ID.Siswa"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   1
         Left            =   7080
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "No.Dokumen"
         Height          =   315
         Index           =   0
         Left            =   7080
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Transaksi_CutiFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim snodokumen As String
Dim stanggal As String
Dim snoidsiswalama As String
Dim snoidsiswa As String
Dim stglmulaicuti As String
Dim stglselesaicuti As String
Dim stglmasukkembali As String
Dim slamacuti As Integer
Dim skeperluancuti As String
Dim skonfirmasists As String
Dim stglkonfirmasi As String
Dim sketkonfirmasi As String
Dim sdokumensts As String
Dim skonfirmasisiswasts As String
Dim smasukkembali As String
Dim skonfirmasitglmasuk As String
Dim saudituser As String
Dim sauditdate As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from vtransaksi_cuti where nodokumen='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
        showData
        
        
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
    sQuery = "Select *  from vtransaksi_cuti order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_cuti where nodokumen >'" & Text1(0).Text & "' order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_cuti where nodokumen<'" & Text1(0).Text & "' order by nodokumen desc limit 1"
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
    sQuery = "Select *  from vtransaksi_cuti order by nodokumen desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
End Sub
Private Function DoSaveData() As Boolean
Dim sQuery2 As String
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        sQuery2 = "Call sp_update_siswa_sts ('" & snoidsiswa & "','" & snoidsiswa & "','" & skonfirmasisiswasts & "')"
        oCon.Execute sQuery
     
        Else
        sQuery = sUpdate
        sQuery2 = "Call sp_update_siswa_sts ('" & snoidsiswalama & "','" & snoidsiswa & "','" & skonfirmasisiswasts & "')"
        oCon.Execute sQuery
        
        End If
        doSaveOgridDetail1
        oCon.Execute sQuery2
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
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "cALL sp_delete_transaksi_cuti('" & snodokumen & "')"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
    Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Dim iDate As Integer
    For iDate = 0 To FlatDatePicker1.Count - 1
    FlatDatePicker1(iDate).value = Now()
    Next
    Option1(0).value = True
    Option2(1).value = True
    Check1(1).value = 0
    Check1(2).value = 0
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler

If istatus = StatusForm.DataBaru Then
    snodokumen = IIf(Text1(0).Text = "", GetDocnum(Transaksi_Cuti, True, parkir), Text1(0).Text)
    Text1(0).Text = snodokumen
Else
    snodokumen = Text1(0).Text
End If
If Option1(0).value = True Then
    sdokumensts = "1"
Else
  sdokumensts = "0"
End If
If Option2(0).value = True Then
    skonfirmasisiswasts = "1"
Else
    skonfirmasisiswasts = "2"
End If

If Check1(1).value = 1 Then
    smasukkembali = "1"
Else
    smasukkembali = "0"
End If
skonfirmasitglmasuk = Format(FlatDatePicker1(5).value, "YYYY-MM-DD")

    snoidsiswa = Text1(1).Text
    stanggal = FlatDatePicker1(0).value
    stglmulaicuti = FlatDatePicker1(1).value
    stglselesaicuti = FlatDatePicker1(2).value
    stglmasukkembali = FlatDatePicker1(3).value
    stglkonfirmasi = FlatDatePicker1(4).value
    slamacuti = ToNumber(Text1(4).Text)
    sketkonfirmasi = Text1(5).Text
     
    sQuery = "update transaksi_cuti "
    sQuery = sQuery & " set "
    sQuery = sQuery & "tanggal='" & Format(stanggal, "yyyy/mm/dd") & " ',"
    sQuery = sQuery & "noidsiswa='" & snoidsiswa & " ',"
    sQuery = sQuery & "tglmulaicuti='" & Format(stglmulaicuti, "yyyy/mm/dd") & " ',"
    sQuery = sQuery & "tglselesaicuti='" & Format(stglselesaicuti, "yyyy/mm/dd") & " ',"
    sQuery = sQuery & "tglmasukkembali='" & Format(stglmasukkembali, "yyyy/mm/dd") & " ',"
    sQuery = sQuery & "lamacuti='" & slamacuti & " ',"
    sQuery = sQuery & "keperluancuti='" & skeperluancuti & " ',"
    sQuery = sQuery & "konfirmasists ='" & skonfirmasists & " ',"
    sQuery = sQuery & "tglkonfirmasi='" & Format(stglkonfirmasi, "yyyy/mm/dd") & " ',"
    sQuery = sQuery & "ketkonfirmasi='" & sketkonfirmasi & " ',"
    sQuery = sQuery & "dokumensts='" & sdokumensts & " ',"
    sQuery = sQuery & "konfirmasisiswasts='" & skonfirmasisiswasts & " ',"
    sQuery = sQuery & "masukkembali='" & smasukkembali & " ',"
    sQuery = sQuery & "konfirmasitglmasuk='" & skonfirmasitglmasuk & " ',"
    sQuery = sQuery & "audituser='" & saudituser & " ',"
    sQuery = sQuery & "auditdate='" & Format(sauditdate, "yyyy/mm/dd") & " '"
    sQuery = sQuery & " where nodokumen= '" & snodokumen & "'"
    sUpdate = sQuery
    
    sQuery = "insert into  transaksi_cuti"
    sQuery = sQuery & "(nodokumen,"
    sQuery = sQuery & "tanggal,"
    sQuery = sQuery & "noidsiswa,"
    sQuery = sQuery & "tglmulaicuti,"
    sQuery = sQuery & "tglselesaicuti,"
    sQuery = sQuery & "tglmasukkembali,"
    sQuery = sQuery & "lamacuti,"
    sQuery = sQuery & "keperluancuti,"
    sQuery = sQuery & "tglkonfirmasi,"
    sQuery = sQuery & "ketkonfirmasi,"
    sQuery = sQuery & "dokumensts,konfirmasisiswasts,masukkembali,konfirmasitglmasuk,"
    sQuery = sQuery & "audituser,"
    sQuery = sQuery & "auditdate)"
    sQuery = sQuery & " values "
    sQuery = sQuery & "("
    sQuery = sQuery & "'" & snodokumen & "',"
    sQuery = sQuery & "'" & Format(stanggal, "yyyy/mm/dd") & "',"
    sQuery = sQuery & "'" & snoidsiswa & "',"
    sQuery = sQuery & "'" & Format(stglmulaicuti, "yyyy/mm/dd") & "',"
    sQuery = sQuery & "'" & Format(stglselesaicuti, "yyyy/mm/dd") & "',"
    sQuery = sQuery & "'" & Format(stglmasukkembali, "yyyy/mm/dd") & "',"
    sQuery = sQuery & "'" & slamacuti & "',"
    sQuery = sQuery & "'" & skeperluancuti & "',"
    sQuery = sQuery & "'" & Format(stglkonfirmasi, "yyyy/mm/dd") & "',"
    sQuery = sQuery & "'" & sketkonfirmasi & "',"
    sQuery = sQuery & "'" & sdokumensts & "',"
    sQuery = sQuery & "'" & skonfirmasisiswasts & "',"
    sQuery = sQuery & "'" & smasukkembali & "',"
    sQuery = sQuery & "'" & skonfirmasitglmasuk & "',"
    sQuery = sQuery & "'" & saudituser & "',"
    sQuery = sQuery & "'" & Format(sauditdate, "yyyy/mm/dd") & "')"
    sInsert = sQuery
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsCuti, ""
    If Not oBrowse.YangDipilih = "" Then
        FindData oBrowse.YangDipilih
    End If
Case 1
    oBrowse.ShowFinder BrowsSiswa, "stssiswa<>'0'"
    If Not oBrowse.YangDipilih = "" Then
        If oFindByQuery("select stssiswa from master_siswa where noidsiswa='" & oBrowse.YangDipilih & "'", parkir) = "2" Then
            MsgBox " Maaf Data Siswa Yang Dipilih Masih Status Siswa Cuti ", vbInformation
            Text1(1).SetFocus
        Else
        
        Text1(1).Text = oBrowse.YangDipilih
        Text1(2).Text = oBrowse.Keterangan
        Text1(3).Text = oFindByQuery("select CONCAT(almtrumah1,IF(almtrumah2='','',CONCAT(',',almtrumah2))) AS alamat from master_siswa where noidsiswa='" & Text1(1).Text & "'", parkir)
        End If
    End If
End Select

Set oBrowse = Nothing
End Sub





Private Sub Check1_Click(Index As Integer)
If Check1(0).value = 1 Then
    Text1(0).Enabled = True
    skonfirmasists = "1"
Else
    skonfirmasists = "0"
    Text1(0).Enabled = False
End If
If Check1(1).value = 1 Then
    FlatDatePicker1(4).Enabled = True
    Text1(5).Enabled = True
Else
    FlatDatePicker1(4).Enabled = False
    Text1(5).Enabled = False
End If

If Check1(1).value = 0 Then
        Frame1(4).Caption = "Belum Ada Konfirmasi"
Else
        Frame1(4).Caption = "Sudah Ada Konfirmasi"
End If
If Check1(2).value = 1 Then
        Option2(0).value = True
        Option2(1).value = False
        FlatDatePicker1(5).Enabled = True
        FlatDatePicker1(5).value = Now()
Else
        Option2(0).value = False
        Option2(1).value = True
        FlatDatePicker1(5).Enabled = False
         FlatDatePicker1(5).value = ""
End If

End Sub

Private Sub FlatDatePicker1_LostFocus(Index As Integer)
Select Case Index
Case 1
    Text1(4) = FlatDatePicker1(2).value - FlatDatePicker1(1).value
  
Case 2
    Text1(4) = FlatDatePicker1(2).value - FlatDatePicker1(1).value
End Select
End Sub

Private Sub FlatDatePicker1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1
    Text1(4) = FlatDatePicker1(2).value - FlatDatePicker1(1).value
  
Case 2
    Text1(4) = FlatDatePicker1(2).value - FlatDatePicker1(1).value
End Select
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Register Cuti Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegisterFormCuti
If Option1(0).value = True Then
        MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = True
    Else
        MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = False
    End If
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
End Sub

Private Sub Form_Load()
oFormatCheckList 1, Me
oFormatOption 2, Me
oSetTanggal
istatus = Normal
cleardata

    FlatDatePicker1(4).Enabled = False
    Text1(5).Enabled = False
    
MoveLast
oFormatOption 1, Me
End Sub

Private Sub showData()
On Error GoTo errhandler

    
    cleardata 'konfirmasisiswasts
    If oRs("dokumensts") = "1" Then
     Option1(0).value = True
     Option1(1).value = False
    Else
     Option1(1).value = True
     Option1(0).value = False
    End If
    If Option1(0).value = True Then
        MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = True
    Else
        MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = False
    End If
    
     
     If oRs("konfirmasisiswasts") = "1" Then
     Option2(0).value = True
     Else
     Option2(1).value = True
     End If
     
    Text1(0).Text = oRs("nodokumen")
    KodeUserAksesTemp = oRs("nodokumen")
    Text1(0).Locked = True
    snoidsiswalama = oRs("noidsiswa")
    Text1(1).Text = oRs("noidsiswa")
    Text1(2).Text = ToText(oRs("nmlengkap"))
    skonfirmasists = oRs("konfirmasists")
    Text1(3).Text = ToText(oRs("alamat"))
    FlatDatePicker1(0).value = oRs("tanggal")
    FlatDatePicker1(1).value = oRs("tglmulaicuti")
    FlatDatePicker1(2).value = oRs("tglselesaicuti")
    FlatDatePicker1(3).value = oRs("tglmasukkembali")
    FlatDatePicker1(4).value = oRs("tglkonfirmasi")
    Text1(4).Text = oRs("lamacuti")
    Text1(5).Text = oRs("ketkonfirmasi")
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    If oRs("konfirmasists") = "0" Then
        Frame1(4).Caption = "Belum Ada Konfirmasi"
    Else
        Frame1(4).Caption = "Sudah Ada Konfirmasi"
    End If
    Check1(1).value = ToNumber(oRs("konfirmasists"))
    
    If oRs("masukkembali") = "1" Then
        Check1(1).value = "1"
    Else
        Check1(1).value = "0"
    End If
    FlatDatePicker1(5).value = oRs("konfirmasitglmasuk")
    
    ShowGrid1 Text1(0).Text, istatus
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

Private Sub oGrid1_Click()
With oGrid1
    Select Case .col
    Case 0
            .EditCell
    End Select
End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1
    
    ShowGrid1 Text1(1).Text, istatus
End Select
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
If Index = 0 Then FindData Text1(0).Text
End Sub

Public Sub oSetTanggal()
Dim i As Integer
For i = 0 To FlatDatePicker1.Count - 1
    FlatDatePicker1(i).CustomFormat = "DD-MM-YYYY"
Next
End Sub

Public Sub ShowGrid1(keynodokumen As String, istatus As StatusForm)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    If istatus = Normal Then
        sQuery = "" & "SELECT b.docentry,a.nokursus,a.tglmulai, b.kodelevelno,b.kodelevel,b.kodegroup,c.keterangan,if(tcd.stscuti='1',-1,0) as stscuti "
        sQuery = sQuery & "FROM  transaksi_cuti_detail1 tcd "
        sQuery = sQuery & "INNER JOIN master_kelas a ON tcd.nokursus=a.nokursus "
        sQuery = sQuery & "INNER JOIN  master_kartu_kelas_detail1 b  ON a.docentry=b.docentry  "
        sQuery = sQuery & "INNER JOIN master_default_pelajaran c ON c.pelajaran=a.pelajaran "
        sQuery = sQuery & "WHERE tcd.nodokumen='" & keynodokumen & "' AND b.aktif='1'  "
        sQuery = sQuery & "GROUP BY b.docentry,a.nokursus, b.kodelevelno,b.kodelevel,b.kodegroup,c.keterangan "
    Else
    
        sQuery = "" & "SELECT b.docentry,a.nokursus,a.tglmulai, b.kodelevelno,b.kodelevel,b.kodegroup,c.keterangan,-1 as stscuti "
        'sQuery = sQuery & "FROM  transaksi_cuti_detail1 tcd "
        sQuery = sQuery & "from  master_kelas a " 'ON tcd.nokursus=a.nokursus "
        sQuery = sQuery & "INNER JOIN  master_kartu_kelas_detail1 b  ON a.docentry=b.docentry  "
        sQuery = sQuery & "INNER JOIN master_default_pelajaran c ON c.pelajaran=a.pelajaran "
        sQuery = sQuery & "WHERE a.noidsiswa ='" & keynodokumen & "' AND b.aktif='1'  "
        sQuery = sQuery & "GROUP BY b.docentry,a.nokursus, b.kodelevelno,b.kodelevel,b.kodegroup,c.keterangan "
    
    End If
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        '.ColHidden(.Cols - 1) = True
        '.Cols = 4
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                '.stscuti(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("stscuti"))
                .TextMatrix(i, 1) = ToText(oRsDetail("nokursus"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("tglmulai"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("keterangan"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("kodelevel"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("kodelevelno"))
                
               
                oRsDetail.MoveNext
            Loop
            '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "ShowGrid1"
End Sub
Public Sub Execution()
On Error GoTo errhandler
Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\CutiFrm.Rpt"

Dim sKriteria As String

sQuery = "SELECT * from vtransaksi_cuti_rpt vtransaksi_cuti_rpt1 where nodokumen='" & Text1(0) & "'"

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

Public Sub doSaveOgridDetail1()
Dim snokursus As String
Dim sstscuti As String
Dim irow As Integer
Dim sQuery As String
With oGrid1
    For irow = 1 To .Rows - 1
        snokursus = .TextMatrix(irow, 1)
        sstscuti = IIf(.TextMatrix(irow, 0) = -1, "1", "0")
        
        If istatus = DataBaru Then
                sQuery = "call sp_insert_transaksi_cuti_detail1('" & snodokumen & "','" & snokursus & "','" & sstscuti & "')"
        Else
                sQuery = "call sp_update_transaksi_cuti_detail1('" & snodokumen & "','" & snokursus & "','" & sstscuti & "')"
        End If
        oExecute sQuery
    Next
End With
End Sub







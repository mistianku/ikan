VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MasukLainFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Transaksi Masuk Lain-Lain Form"
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
      Height          =   1455
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   6600
      Width           =   11775
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   9120
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
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
         Left            =   2280
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   4575
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
         Index           =   10
         Left            =   2280
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Referensi"
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total "
         Height          =   315
         Index           =   10
         Left            =   6960
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total Potongan"
         Height          =   315
         Index           =   9
         Left            =   6960
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total Seb. Potongan"
         Height          =   315
         Index           =   8
         Left            =   6960
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   3375
      Index           =   3
      Left            =   120
      TabIndex        =   28
      Top             =   3240
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   3015
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   5318
         Appearance      =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   4194432
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12632256
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MasukLainFrm.frx":0000
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
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            MouseIcon       =   "MasukLainFrm.frx":0200
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
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2160
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
         Index           =   8
         Left            =   4920
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   6675
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
         Left            =   4920
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   240
         Width           =   6675
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
         Left            =   2280
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   2115
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
         Left            =   2280
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   240
         Width           =   2115
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   22
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MasukLainFrm.frx":021C
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
         Left            =   4440
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MasukLainFrm.frx":0238
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
         Index           =   4
         Left            =   4440
         TabIndex        =   24
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MasukLainFrm.frx":0254
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
         Left            =   2280
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   600
         Width           =   2115
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
         Left            =   4920
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   600
         Width           =   6675
      End
      Begin VB.Label Label1 
         Caption         =   "Ke Gudang"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Harga"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Diskon"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Dokumen"
      Height          =   615
      Index           =   1
      Left            =   9720
      TabIndex        =   12
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   9120
         TabIndex        =   43
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "Auto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8160
         TabIndex        =   42
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
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
         Left            =   2280
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   4635
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
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   4635
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
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   2115
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
         Left            =   9120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   11160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MasukLainFrm.frx":0270
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
         Left            =   4440
         TabIndex        =   10
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MasukLainFrm.frx":028C
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
         Caption         =   "Alamat Supplier"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Supplier"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   2
         Left            =   6960
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Supplier"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No.Dokumen"
         Height          =   315
         Index           =   0
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "MasukLainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm

Dim sdocentry As Double
Dim snodokumen As String
Dim stgldokumen As String
Dim sdokstatus As String
Dim stipetransaksi As String
Dim skodesupplier As String
Dim skodegudang As String
Dim skodeharga As String
Dim skodediskon As String
Dim sketerangan As String
Dim sreferensi As String
Dim stotalsebpotongan As String
Dim stotalpotongan As String
Dim stotalsetpotongan As String

Dim slinenum1 As Integer
Dim skodeproduk1 As String
Dim skodeharga1 As String
Dim skodediskon1 As String
Dim sharga1 As String
Dim sjumlah1 As String
Dim sdiskonpersen1 As String
Dim stotalsebdiskon1 As String
Dim sdiskontotal1 As String
Dim stotalsetdiskon1 As String
Dim skodegudang1 As String
Dim skodegudanglama As String

Dim smodelkwitansi As String
Dim slebar As Integer
Dim stinggi As Integer
Dim stxtpesan As String

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from vtransaksi_masuk_lain where nodokumen='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm
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
    sQuery = "Select *  from vtransaksi_masuk_lain order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_masuk_lain where nodokumen >'" & Text1(0).text & "' order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_masuk_lain where nodokumen<'" & Text1(0).text & "' order by nodokumen desc limit 1"
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
    sQuery = "Select *  from vtransaksi_masuk_lain order by nodokumen desc limit 1 "
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
             FindData Text1(0)
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
        
        If istatus = DataBaru Then
            sdocentry = oFindByQuery("select docentry from transaksi_masuk_lain where nodokumen='" & Text1(0) & "'", DBaseConection.Modul)
        End If
        SaveGrid sdocentry
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
        sQuery = "Delete from transaksi_masuk_lain where nodokumen='" & snodokumen & "'"
        oCon.Execute sQuery
        oCon.Close
        DeleteGrid sdocentry
        DoDeleteData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    If MenuFrm.sAplikasiDemo Then
        If oCekJumlahTrx("transaksi_masuk_lain", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    FlatDatePicker1.value = Now()
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm
    Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    GridModul.ClearGridDetail ogrid
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    snodokumen = ToText(Text1(0).text)
    'sdocentry = oRs("docentry")
    If istatus = StatusForm.DataBaru Then
        snodokumen = ToText(IIf(Text1(0).text = "", GetDocnum(transaksi_masuklain, True, DBaseConection.Modul), Text1(0).text))
        Text1(0).text = snodokumen
    Else
        snodokumen = ToText(Text1(0).text)
    End If
    skodesupplier = ToText(Text1(1).text)
    stgldokumen = Format(FlatDatePicker1.value, "YYYY-MM-DD")
    If Option1(0).value = True Then
        sdokstatus = "1"
    Else
        sdokstatus = "0"
    End If
    
    skodesupplier = ToText(Text1(1))
    skodegudang = ToText(Text1(4))
    skodediskon = ToText(Text1(5))
    skodeharga = ToText(Text1(6))
    sketerangan = ToText(Text1(10))
    sreferensi = ToText(Text1(11))
    stotalsebpotongan = toNumberIndonesia(Text1(12))
    stotalpotongan = toNumberIndonesia(Text1(13))
    stotalsetpotongan = toNumberIndonesia(Text1(14))
    
     
    sUpdate = "update transaksi_masuk_lain set "
    sUpdate = sUpdate & "tgldokumen ='" & stgldokumen & "',"
    sUpdate = sUpdate & "dokstatus ='" & sdokstatus & "',"
    sUpdate = sUpdate & "tipetransaksi ='" & stipetransaksi & "',"
    sUpdate = sUpdate & "kodesupplier ='" & skodesupplier & "',"
    sUpdate = sUpdate & "kodegudang ='" & skodegudang & "',"
    sUpdate = sUpdate & "kodeharga ='" & skodeharga & "',"
    sUpdate = sUpdate & "kodediskon ='" & skodediskon & "',"
    sUpdate = sUpdate & "keterangan ='" & sketerangan & "',"
    sUpdate = sUpdate & "referensi ='" & sreferensi & "',"
    sUpdate = sUpdate & "totalsebpotongan ='" & stotalsebpotongan & "',"
    sUpdate = sUpdate & "totalpotongan ='" & stotalpotongan & "',"
    sUpdate = sUpdate & "totalsetpotongan ='" & stotalsetpotongan & "',"
    sUpdate = sUpdate & "audituser ='" & MenuFrm.sUserID & "',"
    sUpdate = sUpdate & "auditdate ='" & Format(Now(), "YYYY-MM-DD") & "'"
    sUpdate = sUpdate & " where nodokumen= '" & snodokumen & "'"
    
    sInsert = "insert into transaksi_masuk_lain "
    sInsert = sInsert & "("
    sInsert = sInsert & "nodokumen,"
    sInsert = sInsert & "tgldokumen,"
    sInsert = sInsert & "dokstatus,"
    sInsert = sInsert & "tipetransaksi,"
    sInsert = sInsert & "kodesupplier,"
    sInsert = sInsert & "kodegudang,"
    sInsert = sInsert & "kodeharga,"
    sInsert = sInsert & "kodediskon,"
    sInsert = sInsert & "keterangan,"
    sInsert = sInsert & "referensi,"
    sInsert = sInsert & "totalsebpotongan,"
    sInsert = sInsert & "totalpotongan,"
    sInsert = sInsert & "totalsetpotongan,"
    sInsert = sInsert & "audituser,"
    sInsert = sInsert & "auditdate"
    sInsert = sInsert & ")"
    sInsert = sInsert & " values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & snodokumen & "',"
    sInsert = sInsert & "'" & stgldokumen & "',"
    sInsert = sInsert & "'" & sdokstatus & "',"
    sInsert = sInsert & "'" & stipetransaksi & "',"
    sInsert = sInsert & "'" & skodesupplier & "',"
    sInsert = sInsert & "'" & skodegudang & "',"
    sInsert = sInsert & "'" & skodeharga & "',"
    sInsert = sInsert & "'" & skodediskon & "',"
    sInsert = sInsert & "'" & sketerangan & "',"
    sInsert = sInsert & "'" & sreferensi & "',"
    sInsert = sInsert & "'" & stotalsebpotongan & "',"
    sInsert = sInsert & "'" & stotalpotongan & "',"
    sInsert = sInsert & "'" & stotalsetpotongan & "',"
    sInsert = sInsert & "'" & MenuFrm.sUserID & "',"
    sInsert = sInsert & "'" & Format(Now(), "YYYY-MM-DD") & "'"
    sInsert = sInsert & ")"

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
oBrowse.ShowFinder BrowsMasukLain, "", ubDescending, DBaseConection.Modul
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsSupplier, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
        ShowSupplier Text1(1)
        
    End If
Case 2
    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(4) = oBrowse.YangDipilih
        Text1(7) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsDiskon, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(5) = oBrowse.YangDipilih
        Text1(8) = oBrowse.Keterangan
    End If
Case 4
    oBrowse.ShowFinder BrowsHarga, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(6) = oBrowse.YangDipilih
        Text1(9) = oBrowse.Keterangan
    End If
Case 5
    oBrowse.ShowFinder BrowsMasterProduk, "aktif='1'", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        With ogrid
            .TextMatrix(.row, 0) = oBrowse.YangDipilih
            .TextMatrix(.row, 1) = oBrowse.Keterangan
            '.TextMatrix(.Row, 1) = oFindByQuery("select namaproduk from master_produk where kodeproduk='" & .TextMatrix(.Row, 0) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 3) = oFindByQuery("select harga from master_produk_harga where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 11) = 0 'oFindByQuery("select diskon from master_produk_diskon where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodediskon='" & Text1(5) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 5) = ToNumber(.TextMatrix(.row, 3)) * (ToNumber(.TextMatrix(.row, 11)) / 100)
       .Select .row, 2
        End With
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Check1_Click()
If Check1.value = 0 Then
    Text1(0).Enabled = True
Else
    Text1(0).Enabled = False
End If
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Transaksi Masuk Lain-Lain"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasukLainFrm

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
smodelkwitansi = oFindByQuery("select modelkwitansi from master_setting_kwitansi_form", DBaseConection.Modul)
slebar = oFindByQuery("select lebar from master_setting_kwitansi_form", DBaseConection.Modul)
stinggi = oFindByQuery("select tinggi from master_setting_kwitansi_form", DBaseConection.Modul)
stxtpesan = oFindByQuery("select txtpesan tinggi from master_setting_kwitansi_form", DBaseConection.Modul)

oFormatOption 1, Me
cleardata
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    sdocentry = oRs("docentry")
    Text1(0).text = oRs("nodokumen")
    KodeUserAksesTemp = oRs("nodokumen")
    Text1(0).Locked = True
    Text1(1).text = oRs("kodesupplier")
    FlatDatePicker1.value = oRs("tgldokumen")
    If "1" = oRs("dokstatus") Then
    Option1(0).value = True
    Option1(1).value = False
    Else
    Option1(0).value = False
    Option1(1).value = True
    End If
    Text1(1) = oRs("kodesupplier")
        Text1(2) = oRs("namasupplier")
            Text1(3) = oRs("alamat")
    Text1(4) = oRs("kodegudang")
    Text1(5) = oRs("kodediskon")
    Text1(6) = oRs("kodeharga")

    Text1(7) = oRs("namagudang")
    Text1(8) = oRs("namadiskon")
    Text1(9) = ToText(oRs("namaharga"))

    Text1(10) = oRs("keterangan")
    Text1(11) = oRs("referensi")
    Text1(12) = formatRupiah(oRs("totalsebpotongan"))
    Text1(13) = formatRupiah(oRs("totalpotongan"))
    Text1(14) = formatRupiah(oRs("totalsetpotongan"))
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    ShowGrid sdocentry
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub

Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).text = ""
Next
GridModul.ClearGridDetail ogrid
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub



Private Sub ogrid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
BrowseUserID(5).Visible = False
End Sub

Private Sub ogrid_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = ogrid.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, ogrid, istatus
With ogrid
Select Case .col
Case 0, 2
    oRecalculate
End Select
End With

End Sub

Private Sub ogrid_Click()
With ogrid
If .Rows = 1 And Not Text1(1) = "" Then
    AddRow
End If
End With
End Sub

Private Sub ogrid_EnterCell()
With ogrid
    BrowseUserID(5).Visible = False
    Select Case .col
        Case 0
            If .Rows = 1 Then Exit Sub
                            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            SetFinder BrowseUserID(5), ogrid, .col
                            '.EditCell
                             

         Case 2
            .EditCell
            
    End Select
End With
End Sub

Private Sub oGrid_GotFocus()
With ogrid
    BrowseUserID(5).Visible = False
    Select Case .col
        Case 0
            If .Rows = 1 Then Exit Sub
                            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            SetFinder BrowseUserID(5), ogrid, .col
                            '.EditCell
                             

         Case 2
            .EditCell
               
            
    End Select
End With
End Sub

Private Sub ogrid_KeyDown(KeyCode As Integer, Shift As Integer)

With ogrid
If KeyCode = vbKeyDelete Then
   .TextMatrix(.row, 3) = 0
   oRecalculate
End If
MainModule.DoKeyDown KeyCode, istatus


    'If Not ToNumber(.TextMatrix(.row, .Cols - 1)) = 0 Then Exit Sub
    If Not KeyCode = vbKeyInsert Then
           gridDetail_KeyDown KeyCode, 0, ogrid, istatus
           If KeyCode = vbKeyDelete Then Exit Sub
           Select Case .col
           Case 0
                If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                .EditCell
           Case 2
                .EditCell
           End Select
          
           'MsgBox "test"

    Else
        AddRow
        If .col = 0 Then
            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            SetFinder BrowseUserID(5), ogrid, .col
        End If
    End If
End With

End Sub

Private Sub ogrid_KeyDownEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim sKodeQ As String
With ogrid
If Not KeyCode = 13 Then Exit Sub

Select Case col
Case 0
    .Select .row, 0
    sKodeQ = oFindByQuery("SELECT `fget_kodeproduk_from_barcode`('" & .TextMatrix(.row, 0) & "')", DBaseConection.Modul)
    If sKodeQ = "" Then
         MsgBox "Master Produk Tidak Ditemukan", vbInformation
        .Select .row, 0
    Else
        .TextMatrix(.row, 0) = sKodeQ
         .TextMatrix(.row, 1) = oFindByQuery("select namaproduk from master_produk where kodeproduk='" & sKodeQ & "'", DBaseConection.Modul)
         .TextMatrix(.row, 3) = oFindByQuery("select harga from master_produk_harga where kodeproduk='" & sKodeQ & "' and kodeharga='" & Text1(6) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 11) = 0 ' oFindByQuery("select diskon from master_produk_diskon where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodediskon='" & Text1(5) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 5) = ToNumber(.TextMatrix(.row, 3)) * (ToNumber(.TextMatrix(.row, 11)) / 100)
       .Select .row, 2
       .EditCell
    End If
'
'    If oFindByQuery("select namaproduk from master_produk where kodeproduk='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul) = "" Then
'        MsgBox "Master Produk Tidak Ditemukan", vbInformation
'        .Select .row, 0
'    Else
'         .TextMatrix(.row, 1) = oFindByQuery("select namaproduk from master_produk where kodeproduk='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul)
'         .TextMatrix(.row, 3) = oFindByQuery("select harga from master_produk_harga where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "'", DBaseConection.Modul)
'         .TextMatrix(.row, 11) = 0 ' oFindByQuery("select diskon from master_produk_diskon where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodediskon='" & Text1(5) & "'", DBaseConection.Modul)
'         .TextMatrix(.row, 5) = ToNumber(.TextMatrix(.row, 3)) * (ToNumber(.TextMatrix(.row, 11)) / 100)
'       .Select .row, 2
'       .EditCell
'    End If

Case 2
    If .Rows - 1 = row Then
        AddRow
        '.Select .row, 0
    Else
        .Select .row + 1, 0
    End If
End Select


End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 4
    oUpdateKodeGudang Text1(Index), ogrid
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
If Index = 0 Then FindData Text1(0).text
End Sub

Public Sub ShowGrid(sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    sKondisi = " Where docentry=" & sdocentry & " Order by linenum asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "SELECT * FROM vtransaksi_masuk_lain_detail1  "
    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

        '.COLWIDTH(1) = .Width - (.COLWIDTH(0) + .COLWIDTH(2) + .COLWIDTH(3)) - 100
        GridModul.ClearGridDetail ogrid
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodeproduk"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namaproduk"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("jumlah"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("harga"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("totalsebdiskon"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("diskontotal"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("totalsetdiskon"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("jumlah"))
                .TextMatrix(i, 8) = RTrim(oRsDetail("kodegudang"))
                .TextMatrix(i, 9) = RTrim(oRsDetail("kodediskon"))
                .TextMatrix(i, 10) = RTrim(oRsDetail("kodeharga"))
                .TextMatrix(i, 11) = RTrim(oRsDetail("diskonpersen"))
                .TextMatrix(i, .Cols - 4) = RTrim(oRsDetail("kodegudang"))
                .TextMatrix(i, .Cols - 3) = sdocentry
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("linenum"))
                .TextMatrix(i, .Cols - 1) = 0
                oRsDetail.MoveNext
            Loop
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub ShowSupplier(skodesupplier As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    sKondisi = " Where kodesupplier='" & skodesupplier & "' limit 1 "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "SELECT * FROM vmaster_supplier  "
    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    If Not oRsDetail.EOF Then
        Text1(3) = oRsDetail("txtalamat")
        Text1(4) = oRsDetail("kodegudang")
        Text1(5) = oRsDetail("kodediskon")
        Text1(6) = oRsDetail("kodeharga")
        Text1(7) = oRsDetail("namagudang")
        Text1(8) = oRsDetail("namadiskon")
        Text1(9) = oRsDetail("namaharga")
    End If
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub oRecalculate()
Dim irow As Integer
Dim sttlawal As Double
Dim sttlpot As Double
Dim sttlstlpot As Double
With ogrid
    
    For irow = 1 To .Rows - 1
        .TextMatrix(irow, 4) = ToNumber(.TextMatrix(irow, 2)) * ToNumber(.TextMatrix(irow, 3))
        '.TextMatrix(irow, 5) = ToNumber(.TextMatrix(irow, 2)) * (ToNumber(.TextMatrix(irow, 3)) * ToNumber(.TextMatrix(irow, 9)))
        .TextMatrix(irow, 6) = ToNumber(.TextMatrix(irow, 4)) - ToNumber(.TextMatrix(irow, 5))
        sttlawal = sttlawal + ToNumber(.TextMatrix(irow, 4))
        sttlpot = sttlpot + ToNumber(.TextMatrix(irow, 5))
        sttlstlpot = sttlstlpot + ToNumber(.TextMatrix(irow, 6))
    Next
        Text1(12) = Format(sttlawal, "###,###,###.#0")
        Text1(13) = Format(sttlpot, "###,###,###.#0")
        Text1(14) = Format(sttlstlpot, "###,###,###.#0")
End With
End Sub
Public Sub AddRow()
With ogrid
If .TextMatrix(.row, 0) = "" Then Exit Sub
               
            .Rows = .Rows + 1
            .Select .Rows - 1, 0
            .Cell(flexcpFontBold, .row, 0, , .Cols - 1) = vbNormal
            '.EditCell

        .TextMatrix(.row, 2) = 1
        .TextMatrix(.row, 5) = 0
        .TextMatrix(.row, 6) = 0
        .TextMatrix(.row, 7) = 0
        .TextMatrix(.row, 8) = Text1(4)
        .TextMatrix(.row, 9) = Text1(5)
        .TextMatrix(.row, 10) = Text1(6)
        .TextMatrix(.row, 11) = 0
        If istatus = DataBaru Then
            .TextMatrix(.row, .Cols - 3) = "0"
            If .row = 1 Then
                .TextMatrix(.row, .Cols - 2) = 1
            Else
                .TextMatrix(.row, .Cols - 2) = ToNumber(.TextMatrix(.row - 1, .Cols - 2)) + 1
            End If
        Else
            .TextMatrix(.row, .Cols - 3) = sdocentry
            If .row = 1 Then
                .TextMatrix(.row, .Cols - 2) = 1
            Else
                .TextMatrix(.row, .Cols - 2) = ToNumber(.TextMatrix(.row - 1, .Cols - 2)) + 1
            End If
        End If
        .TextMatrix(.row, .Cols - 1) = "1"
End With
End Sub
Public Sub SaveGrid(sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

            Dim i As Double
            Dim sjumlahseb As Double
            For i = 1 To .Rows - 1
            
                skodeproduk1 = .TextMatrix(i, 0)
                sjumlah1 = toNumberIndonesia(.TextMatrix(i, 2))
                sharga1 = toNumberIndonesia(.TextMatrix(i, 3))
                stotalsebdiskon1 = toNumberIndonesia(.TextMatrix(i, 4))
                sdiskontotal1 = toNumberIndonesia(.TextMatrix(i, 5))
                stotalsetdiskon1 = toNumberIndonesia(.TextMatrix(i, 6))
                sjumlahseb = toNumberIndonesia(.TextMatrix(i, 7))
                skodegudang1 = .TextMatrix(i, 8)
                skodediskon1 = .TextMatrix(i, 9)
                skodeharga1 = .TextMatrix(i, 10)
                sdiskonpersen1 = .TextMatrix(i, 11)
                skodegudanglama = .TextMatrix(i, .Cols - 4)
                slinenum1 = ToNumber(.TextMatrix(i, .Cols - 2))
                Select Case ToNumber(.TextMatrix(i, .Cols - 1))
                Case 1 And Not skodeproduk1 = ""
                        sQuery = "insert into transaksi_masuk_lain_detail1 "
                        sQuery = sQuery & "("
                        sQuery = sQuery & "docentry,"
                        sQuery = sQuery & "linenum,"
                        sQuery = sQuery & "kodeproduk,"
                        sQuery = sQuery & "kodeharga,"
                        sQuery = sQuery & "kodediskon,"
                        sQuery = sQuery & "harga,"
                        sQuery = sQuery & "jumlah,"
                        sQuery = sQuery & "diskonpersen,"
                        sQuery = sQuery & "totalsebdiskon,"
                        sQuery = sQuery & "diskontotal,"
                        sQuery = sQuery & "totalsetdiskon,"
                        sQuery = sQuery & "kodegudang,"
                        sQuery = sQuery & "audituser,"
                        sQuery = sQuery & "auditdate"
                        sQuery = sQuery & ")"
                        sQuery = sQuery & " values "
                        sQuery = sQuery & "("
                        sQuery = sQuery & "'" & sdocentry & "',"
                        sQuery = sQuery & "'" & slinenum1 & "',"
                        sQuery = sQuery & "'" & skodeproduk1 & "',"
                        sQuery = sQuery & "'" & skodeharga1 & "',"
                        sQuery = sQuery & "'" & skodediskon1 & "',"
                        sQuery = sQuery & "'" & sharga1 & "',"
                        sQuery = sQuery & "'" & sjumlah1 & "',"
                        sQuery = sQuery & "'" & sdiskonpersen1 & "',"
                        sQuery = sQuery & "'" & stotalsebdiskon1 & "',"
                        sQuery = sQuery & "'" & sdiskontotal1 & "',"
                        sQuery = sQuery & "'" & stotalsetdiskon1 & "',"
                        sQuery = sQuery & "'" & skodegudang1 & "',"
                        sQuery = sQuery & "'" & MenuFrm.sUserID & "',"
                        sQuery = sQuery & "'" & Format(Now(), "YYYY-MM-DD") & "'"
                        sQuery = sQuery & ")"
                        oKon.Execute sQuery
                        oKon.Execute "update master_inventori set stock=stock+" & sjumlah1 & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
                Case 2
                        oKon.Execute "update master_inventori set stock=stock-" & sjumlahseb & " where kodegudang='" & skodegudanglama & "' and kodeproduk='" & skodeproduk1 & "'"
                        
                        sQuery = "update  transaksi_masuk_lain_detail1 "
                        sQuery = sQuery & " set "
                        sQuery = sQuery & "kodeproduk=  '" & skodeproduk1 & "',"
                        sQuery = sQuery & "kodeharga=  '" & skodeharga1 & "',"
                        sQuery = sQuery & "kodediskon=  '" & skodediskon1 & "',"
                        sQuery = sQuery & "harga=  '" & sharga1 & "',"
                        sQuery = sQuery & "jumlah=  '" & sjumlah1 & "',"
                        sQuery = sQuery & "diskonpersen=  '" & sdiskonpersen1 & "',"
                        sQuery = sQuery & "totalsebdiskon=  '" & stotalsebdiskon1 & "',"
                        sQuery = sQuery & "diskontotal=  '" & sdiskontotal1 & "',"
                        sQuery = sQuery & "totalsetdiskon=  '" & stotalsetdiskon1 & "',"
                        sQuery = sQuery & "kodegudang=  '" & skodegudang1 & "',"
                        sQuery = sQuery & "audituser=  '" & MenuFrm.sUserID & "',"
                        sQuery = sQuery & "auditdate=  '" & Format(Now(), "YYYY-MM-DD") & "'"
                        sQuery = sQuery & " where docentry=  '" & sdocentry & "' "
                        sQuery = sQuery & " and linenum=  '" & slinenum1 & "'"

                        oKon.Execute sQuery
                        oKon.Execute "update master_inventori set stock=stock+" & sjumlah1 & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
                Case 3
                        oKon.Execute "update master_inventori set stock=stock-" & sjumlahseb & " where kodegudang='" & skodegudanglama & "' and kodeproduk='" & skodeproduk1 & "'"
                        oKon.Execute "delete from transaksi_masuk_lain_detail1 where docentry='" & sdocentry & "' and linenum='" & slinenum1 & "'"
                End Select
            Next

        'End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub DeleteGrid(sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

            Dim i As Double
            Dim sjumlahseb As Double
            For i = 1 To .Rows - 1
            
                skodeproduk1 = .TextMatrix(i, 0)
                sjumlah1 = ToNumber(.TextMatrix(i, 2))
                sharga1 = ToNumber(.TextMatrix(i, 3))
                stotalsebdiskon1 = ToNumber(.TextMatrix(i, 4))
                sdiskontotal1 = ToNumber(.TextMatrix(i, 5))
                stotalsetdiskon1 = ToNumber(.TextMatrix(i, 6))
                sjumlahseb = ToNumber(.TextMatrix(i, 7))
                skodegudang1 = .TextMatrix(i, 8)
                skodediskon1 = .TextMatrix(i, 9)
                skodeharga1 = .TextMatrix(i, 10)
                sdiskonpersen1 = .TextMatrix(i, 11)
                slinenum1 = ToNumber(.TextMatrix(i, .Cols - 2))
                
                        oKon.Execute "update master_inventori set stock=stock-" & sjumlahseb & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
                        oKon.Execute "delete from transaksi_masuk_lain_detail1 where docentry='" & sdocentry & "' and linenum='" & slinenum1 & "'"

            Next

        'End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Delete Detail Data"
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



'stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
'stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")

'If Text1(0).Text = "" Then
'    scustomercodefr = oFindByQuery("select custmrcode from mst_customer order by custmrcode asc limit 1 ", DBaseConection.Modul)
'Else
'    scustomercodefr = Text1(0).Text
'End If
'If Text1(2).Text = "" Then
'    scustomercodeto = oFindByQuery("select custmrcode from mst_customer order by custmrcode desc limit 1 ", DBaseConection.Modul)
'Else
'    scustomercodeto = Text1(2).Text
'End If

Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

sQuery = "call sp_transaksi_masuk_lain_form('"
sQuery = sQuery & Text1(0) & "','"
sQuery = sQuery & Text1(0) & "',"



If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If
With arMasukLainFrm
    .lblHeaderTrx = "MASUK LAIN-LAIN"
'    .lblCompany1 = MenuFrm.txtHeader(0)
'    .lblCompany2 = MenuFrm.txtHeader(1)
'    .lblCompany3 = MenuFrm.txtHeader(2)
    
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery & "0)"
    
'    .lblkode.Caption = "Kode Custmr"
'    .lblketerangan.Caption = "Nama Customer"
'    .txtkodeproduk.DataField = "custmrcode"
'    .txtproductname.DataField = "custmrname"
'    .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = 16820
'    .PageSettings.PaperWidth = 11904
    .Show
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Transaksi Masuk Lain-lain"
End Sub

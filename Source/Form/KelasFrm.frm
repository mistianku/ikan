VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form KelasFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   7335
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
   ScaleHeight     =   7335
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Dokumen"
      Enabled         =   0   'False
      Height          =   735
      Index           =   6
      Left            =   8520
      TabIndex        =   66
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   67
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Penggolongan"
      Height          =   735
      Index           =   4
      Left            =   4320
      TabIndex        =   62
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Baru"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Caption         =   "Pindah Masuk"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   64
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Pel.Tambahan"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   63
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informasi Data Kursus"
      Height          =   1455
      Index           =   7
      Left            =   120
      TabIndex        =   43
      Top             =   840
      Width           =   11055
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
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   600
         Width           =   3615
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
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   240
         Width           =   3135
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
         Left            =   8160
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Manual"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   7080
         TabIndex        =   48
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
         Height          =   285
         Index           =   3
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   960
         Width           =   3615
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
         Left            =   8160
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   960
         Width           =   2055
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   8160
         TabIndex        =   44
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
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   10200
         TabIndex        =   45
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "KelasFrm.frx":0000
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
         Left            =   5400
         TabIndex        =   52
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "KelasFrm.frx":001C
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
         Index           =   2
         Left            =   10320
         TabIndex        =   53
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "KelasFrm.frx":0038
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
         Index           =   16
         Left            =   2280
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "No.ID Siswa"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Mulai "
         Height          =   315
         Index           =   1
         Left            =   6000
         TabIndex        =   59
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No. Kursus"
         Height          =   315
         Index           =   0
         Left            =   6000
         TabIndex        =   58
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Pembimbing"
         Height          =   315
         Index           =   2
         Left            =   6000
         TabIndex        =   57
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Reff.No.Pendaftaran"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kelas"
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Matematika"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Inggris EE"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Inggris EFL"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
   End
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   2400
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      MouseIcon       =   "KelasFrm.frx":0054
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Buat Kartu Belajar"
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   12000
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   8280
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Utama"
            Key             =   "mnUtama"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Data Sebelumnya"
            Key             =   "mnDataSebelumnya"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   5295
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   11055
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   8400
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   4320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   7575
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   6360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   825
      End
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   3255
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   10095
         _cx             =   17806
         _cy             =   5741
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"KelasFrm.frx":0070
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
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   34
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         MouseIcon       =   "KelasFrm.frx":016B
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
         Left            =   7200
         TabIndex        =   35
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         MouseIcon       =   "KelasFrm.frx":0187
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
      Begin VB.Label lblgrouppelajaran 
         Alignment       =   2  'Center
         Caption         =   "lblgrouppelajaran"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   1320
         Width           =   10095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kelompok"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   16
         Left            =   8400
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Titik Pangkal"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   6360
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jawaban Yang Benar"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   4320
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Waktu"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   2280
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jenis"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hasil Tes Penempatan / Tes Penyelesaian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   5175
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   11055
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Pelajaran Sekolah"
         Height          =   975
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   3975
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Baik"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sedang"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Buruk"
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Untuk Siswa Pindah Masuk "
         Height          =   2655
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   5895
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
            Left            =   3000
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1080
            Width           =   2655
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
            Left            =   3000
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   1440
            Width           =   2655
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
            Left            =   3000
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   1800
            Width           =   2655
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
            Left            =   3000
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   2160
            Width           =   2655
         End
         Begin NeoCalendarII.DatePicker FlatDatePicker1 
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   36
            Top             =   360
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
         Begin NeoCalendarII.DatePicker FlatDatePicker1 
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   37
            Top             =   720
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
            Caption         =   "Tanggal Masuk"
            Height          =   315
            Index           =   17
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Tanggal Keluar"
            Height          =   315
            Index           =   18
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "No.Kursus Sebelumnya"
            Height          =   315
            Index           =   19
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Pembimbing Sebelumnya"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Titik Pangkal Sebelumnya"
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            Caption         =   "Kemajuan Terakhir Sebelumnya"
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   2775
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
         Index           =   14
         Left            =   4200
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   840
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Catatan Khusus"
         Height          =   315
         Index           =   11
         Left            =   4200
         TabIndex        =   29
         Top             =   480
         Width           =   6015
      End
   End
End
Attribute VB_Name = "KelasFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sdokstatus As String
Dim sreffnosebelumnya As String
Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Dim skodegroup As String
Dim skodelevel As String
Dim sUpdateDetail1 As String
Dim sInsertDetail1 As String

Dim skodelevellama As String
Dim sdocentry As Integer
Dim snokursus  As String
Dim stglmulai  As String
Dim snoidsiswa As String
Dim snamapembimbing As String
Dim spenggolongan  As String
Dim spelajaran As String
Dim slevelno  As Integer
Dim swaktupengerjaan  As String
Dim sjawabanbenar  As String
Dim stitikpangkal  As String
Dim skelompok  As String
Dim snilaisekolah  As String
Dim scatatan  As String
Dim ssptglmasuk As String
Dim ssptglkeluar  As String
Dim sspnoidsiswa  As String
Dim sspnamapembimbing  As String
Dim ssptitikpangkal As String
Dim sspkemajuanterakhir As String
Dim sauditdate As String
Dim saudituser As String
Dim sstskelas As String

Dim stgltest As String
Dim skodelevelno As String
Dim skodelevelnodetail As Integer
Dim saktif As String
Dim sstatus As String
Dim sreffno As String

Dim syop As Integer
Dim smop As Integer
Dim sbasedocentry  As Integer
Dim snokwitansi As String
Dim sstsbayar  As String
Dim sjnsbayar  As String
Dim snilaibayar As String
Dim stglbayar As String
Dim sobjtype  As String



Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from vmaster_kelas where nokursus='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
'        istatus = Normal
'        MenuFrm.SetToolbar istatus
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas
    End If
    'oSetToolBar Normal, sstskelas
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select *  from vmaster_kelas order by nokursus asc limit 1"
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
    sQuery = "Select  *  from vmaster_kelas where nokursus >'" & Text1(0).Text & "' order by nokursus asc limit 1"
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
    sQuery = "Select  *  from vmaster_kelas where nokursus<'" & Text1(0).Text & "' order by nokursus desc limit 1"
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
    sQuery = "Select *  from vmaster_kelas order by nokursus desc limit 1 "
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
    If Text1(6) = "" Then
        MsgBox "Level Belum Di Isi ", vbInformation
        oGrid1.Select oGrid1.row, 5
        oGrid1.EditCell
        Exit Sub
    End If
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             FindData Text1(0)
             
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
        oSaveDetail1
        oCon.Execute "update transaksi_pendaftaran set dokstatuskelas='C' where nopendaftaran='" & sreffno & "'"
        If oFindByQuery("select count(*) from master_pembimbing where namapembimbing='" & snamapembimbing & "'", parkir) = 0 And Text1(4) <> "" Then
            oSimpanMasterPembimbing
        End If
        
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
        sQuery = "call spDelete_master_kelas('" & sdocentry & "')"
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
        If oCekJumlahTrx("master_kelas", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    Text1(1).Enabled = True
    BrowseUserID(1).Enabled = True
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas
    Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    FlatDatePicker1(0).value = Now()
    FlatDatePicker1(1).value = Now()
    FlatDatePicker1(2).value = Now()
    Command1.Enabled = False
 
    sdocentry = 0
    
        Option2(0).value = True
        Command1.Enabled = True
        oGrid1.Enabled = True
        Frame1(3).Enabled = True
        BrowseUserID(4).Enabled = True
        Text1(5).Enabled = True
        Text1(6).Enabled = True
        
        skodegroup = "02"
    If Check1(0).value = 1 Then
       Text1(0).Enabled = True
    Else
        Text1(0).Enabled = False
    End If
    
End Sub
Public Sub Undo()
    istatus = Normal
    FindData KodeUserAksesTemp
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
              
snokursus = Text1(0).Text
stglmulai = Format(FlatDatePicker1(0).value, "YYYY/MM/DD")
snoidsiswa = Text1(1).Text
snamapembimbing = Text1(4).Text
spenggolongan = IIf(Option3(0).value = True, "1", IIf(Option3(1).value = True, "2", "3"))
spelajaran = IIf(Option4(0).value = True, "1", IIf(Option4(1).value = True, "2", "3"))
skodelevel = Text1(5).Text
slevelno = ToNumber(Text1(6).Text)
swaktupengerjaan = Text1(7).Text
sjawabanbenar = Text1(8).Text
stitikpangkal = Text1(9).Text
skelompok = Text1(15).Text
snilaisekolah = IIf(Option1(0).value = True, "1", IIf(Option1(1).value = True, "2", "3"))
scatatan = Text1(14).Text
ssptglmasuk = Format(FlatDatePicker1(1).value, "YYYY/MM/DD")
ssptglkeluar = Format(FlatDatePicker1(2).value, "YYYY/MM/DD")
sspnoidsiswa = Text1(10).Text
sspnamapembimbing = Text1(11).Text
ssptitikpangkal = Text1(12).Text
sspkemajuanterakhir = Text1(13).Text
sreffno = Text1(16).Text
sauditdate = Format(Now(), "YYYY/MM/DD")
saudituser = MenuFrm.sUserID

If istatus = DataBaru Then
    If Check1(0).value = 1 And Text1(0).Text <> "" Then
        snokursus = Text1(0)
    Else
        snokursus = GetDocnum(transaksi_kelas, True, parkir)
        Text1(0).Text = snokursus
    End If
Else
End If
sUpdate = "update master_kelas  set "
sUpdate = sUpdate & ""
sUpdate = sUpdate & "tglmulai='" & stglmulai & "',"
sUpdate = sUpdate & "reffno='" & sreffno & "',"
sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "',"
sUpdate = sUpdate & "namapembimbing='" & snamapembimbing & "',"
sUpdate = sUpdate & "penggolongan='" & spenggolongan & "',"
sUpdate = sUpdate & "pelajaran='" & spelajaran & "',"
sUpdate = sUpdate & "kodegroup='" & skodegroup & "',"
sUpdate = sUpdate & "kodelevel='" & skodelevel & "',"
sUpdate = sUpdate & "levelno='" & slevelno & "',"
sUpdate = sUpdate & "waktupengerjaan='" & swaktupengerjaan & "',"
sUpdate = sUpdate & "jawabanbenar='" & sjawabanbenar & "',"
sUpdate = sUpdate & "titikpangkal='" & stitikpangkal & "',"
sUpdate = sUpdate & "kelompok='" & skelompok & "',"
sUpdate = sUpdate & "nilaisekolah='" & snilaisekolah & "',"
sUpdate = sUpdate & "catatan='" & scatatan & "',"
sUpdate = sUpdate & "sptglmasuk='" & ssptglmasuk & "',"
sUpdate = sUpdate & "sptglkeluar='" & ssptglkeluar & "',"
sUpdate = sUpdate & "spnoidsiswa='" & sspnoidsiswa & "',"
sUpdate = sUpdate & "spnamapembimbing='" & sspnamapembimbing & "',"
sUpdate = sUpdate & "sptitikpangkal='" & ssptitikpangkal & "',"
sUpdate = sUpdate & "spkemajuanterakhir='" & sspkemajuanterakhir & "',"
sUpdate = sUpdate & "auditdate='" & sauditdate & "',"
sUpdate = sUpdate & "audituser='" & saudituser & "'"
sUpdate = sUpdate & " where nokursus='" & snokursus & "'"

    
    sInsert = "insert into master_kelas "
    sInsert = sInsert & "("
    sInsert = sInsert & "nokursus,"
    sInsert = sInsert & "tglmulai,"
    sInsert = sInsert & "reffno,"
    sInsert = sInsert & "noidsiswa,"
    sInsert = sInsert & "namapembimbing,"
    sInsert = sInsert & "penggolongan,"
    sInsert = sInsert & "pelajaran,"
    sInsert = sInsert & "kodegroup,"
    sInsert = sInsert & "kodelevel,"
    sInsert = sInsert & "levelno,"
    sInsert = sInsert & "waktupengerjaan,"
    sInsert = sInsert & "jawabanbenar,"
    sInsert = sInsert & "titikpangkal,"
    sInsert = sInsert & "kelompok,"
    sInsert = sInsert & "nilaisekolah,"
    sInsert = sInsert & "catatan,"
    sInsert = sInsert & "sptglmasuk,"
    sInsert = sInsert & "sptglkeluar,"
    sInsert = sInsert & "spnoidsiswa,"
    sInsert = sInsert & "spnamapembimbing,"
    sInsert = sInsert & "sptitikpangkal,"
    sInsert = sInsert & "spkemajuanterakhir,"
    sInsert = sInsert & "auditdate,"
    sInsert = sInsert & "audituser"
    sInsert = sInsert & ") "
    sInsert = sInsert & "values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & snokursus & "',"
    sInsert = sInsert & "'" & stglmulai & "',"
    sInsert = sInsert & "'" & sreffno & "',"
    sInsert = sInsert & "'" & snoidsiswa & "',"
    sInsert = sInsert & "'" & snamapembimbing & "',"
    sInsert = sInsert & "'" & spenggolongan & "',"
    sInsert = sInsert & "'" & spelajaran & "',"
    sInsert = sInsert & "'" & skodegroup & "',"
    sInsert = sInsert & "'" & skodelevel & "',"
    sInsert = sInsert & "'" & slevelno & "',"
    sInsert = sInsert & "'" & swaktupengerjaan & "',"
    sInsert = sInsert & "'" & sjawabanbenar & "',"
    sInsert = sInsert & "'" & stitikpangkal & "',"
    sInsert = sInsert & "'" & skelompok & "',"
    sInsert = sInsert & "'" & snilaisekolah & "',"
    sInsert = sInsert & "'" & scatatan & "',"
    sInsert = sInsert & "'" & ssptglmasuk & "',"
    sInsert = sInsert & "'" & ssptglkeluar & "',"
    sInsert = sInsert & "'" & sspnoidsiswa & "',"
    sInsert = sInsert & "'" & sspnamapembimbing & "',"
    sInsert = sInsert & "'" & ssptitikpangkal & "',"
    sInsert = sInsert & "'" & sspkemajuanterakhir & "',"
    sInsert = sInsert & "'" & sauditdate & "',"
    sInsert = sInsert & "'" & saudituser & "')"

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsKelas, "'1'='1'"
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsSiswa, "stssiswa<>'0'"
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
        Text1(3) = oFindByQuery("select almtrumah1 from master_siswa where noidsiswa='" & Text1(1) & "'", parkir)
        'oTampildrReff Text1(16)
    End If
Case 2
    oBrowse.ShowFinder BrowsPembimbing, ""
    If Not oBrowse.YangDipilih = "" Then
        'Text1(1) = oBrowse.YangDipilih
        Text1(4) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsJenisMateri, ""
    If Not oBrowse.YangDipilih = "" Then
        'Text1(1) = oBrowse.YangDipilih
        Text1(9) = oBrowse.YangDipilih
    End If
    
Case 4
    oBrowse.ShowFinder BrowsPelajaranLevel, "kodegroup='" & skodegroup & "'"
    If Not oBrowse.YangDipilih = "" Then
        'Text1(1) = oBrowse.YangDipilih
        Text1(5) = oBrowse.YangDipilih
    End If
End Select

Set oBrowse = Nothing
End Sub

Private Sub FlatTab1_Change(TabKey As String)

End Sub

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    If Check1(0).value = 1 Then
       Text1(0).Enabled = True
    Else
        Text1(0).Enabled = False
    End If
End Select
End Sub

Private Sub Command1_Click()
If MsgBox("Proses Buat Kartu Belajar Dilanjutkan", vbInformation + vbYesNo) = vbYes Then
    oCreateKartuKelas sdocentry
End If
    
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Register No Kursus"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnRegistrasiKelas


BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(16).Top
BrowseUserID(1).Height = Text1(16).Height
BrowseUserID(1).Left = Text1(16).Left + Text1(16).Width

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width

BrowseUserID(3).Top = Text1(9).Top
BrowseUserID(3).Height = Text1(9).Height
BrowseUserID(3).Left = Text1(9).Left + Text1(9).Width

BrowseUserID(4).Top = Text1(5).Top
BrowseUserID(4).Height = Text1(5).Height
BrowseUserID(4).Left = Text1(5).Left + Text1(5).Width
End Sub

Private Sub Form_Load()

oFormatOption 4, Me
oFormatCheckList 1, Me
cleardata
istatus = Normal
MoveLast

Frame1(1).Left = Frame1(0).Left
Frame1(1).Top = Frame1(0).Top
Frame1(1).Width = Frame1(0).Width


End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
        If Check1(0).value = 1 Then
       Text1(0).Enabled = True
    Else
        Text1(0).Enabled = False
    End If
    
    sreffnosebelumnya = ToText(oRs("reffno"))
    sreffno = ToText(oRs("reffno"))
    sdocentry = ToText(oRs("docentry"))
    sstskelas = ToText(oRs("stskelas"))
    sdokstatus = ToText(oRs("dokstatus"))
    snokursus = ToText(oRs("nokursus"))
    FlatDatePicker1(0).value = ToText(oRs("tglmulai"))
    
    If sdokstatus = "1" Then
        Option2(0).value = True
        Command1.Enabled = True
        'ogrid1.Enabled = True
        Frame1(3).Enabled = True
        BrowseUserID(1).Enabled = True
        Text1(1).Enabled = True
        BrowseUserID(4).Enabled = True
        Text1(5).Enabled = True
        Text1(6).Enabled = True
    Else
        Option2(1).value = True
        Command1.Enabled = False
        'ogrid1.Enabled = False
        Frame1(3).Enabled = False
        BrowseUserID(1).Enabled = False
        Text1(1).Enabled = False
        BrowseUserID(4).Enabled = False
        Text1(5).Enabled = False
        Text1(6).Enabled = False
    End If
    
    Text1(16).Text = ToText(oRs("reffno"))
    Text1(0).Text = ToText(oRs("nokursus"))
    KodeUserAksesTemp = ToText(oRs("nokursus"))
    Text1(0).Locked = True
    Text1(1).Text = ToText(oRs("noidsiswa"))
    Text1(2).Text = ToText(oRs("nmlengkap"))
    Text1(3).Text = ToText(oRs("almtrumah1"))
    Text1(4).Text = ToText(oRs("namapembimbing"))

    Text1(5).Text = ToText(oRs("kodelevel"))
    skodelevel = ToText(oRs("kodelevel"))
    skodelevellama = ToText(oRs("kodelevel"))
    Text1(6).Text = oRs("levelno")
    Text1(7).Text = ToText(oRs("waktupengerjaan"))
    Text1(8).Text = ToText(oRs("jawabanbenar"))
    Text1(9).Text = ToText(oRs("titikpangkal"))
    Text1(15).Text = ToText(oRs("kelompok"))
    
    Text1(14).Text = ToText(oRs("catatan"))
    FlatDatePicker1(1).value = oRs("sptglmasuk")
    FlatDatePicker1(2).value = oRs("sptglkeluar")
    Text1(10).Text = ToText(oRs("spnoidsiswa"))
    Text1(11).Text = ToText(oRs("spnamapembimbing"))
    Text1(12).Text = ToText(oRs("sptitikpangkal"))
    Text1(13).Text = ToText(oRs("spkemajuanterakhir"))

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    Select Case ToText(oRs("nilaisekolah"))
    Case "1"
        Option1(0).value = True
    Case "2"
        Option1(1).value = True
    Case Else
        Option1(2).value = True
    End Select
    
    Select Case ToText(oRs("penggolongan"))
    Case "1"
        Option3(0).value = True
    Case "2"
        Option3(1).value = True
    Case Else
        Option3(2).value = True
    End Select
    Select Case ToText(oRs("pelajaran"))
    Case "1"
        Option4(0).value = True
    Case "2"
        Option4(1).value = True
    Case Else
        Option4(2).value = True
    End Select


    'Label2.Caption = IIf(ToText(oRs("penggolongan")) = "1", "Baru", IIf(ToText(oRs("penggolongan")) = "2", "Pindah Masuk", "Pelajaran Tambahan"))
    'Label3.Caption = IIf(ToText(oRs("pelajaran")) = "1", "Matematika", IIf(ToText(oRs("pelajaran")) = "2", "Bahasa Inggris EE", "Bahasa Inggris EFL"))
        
    skodegroup = oFindByQuery("select kodegroup from master_default_pelajaran where pelajaran='" & ToText(oRs("pelajaran")) & "'", parkir)
    ShowGrid1 skodegroup, Trim(Text1(5))
    'oSetToolBar Normal, sdokstatus
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

Private Sub ogrid1_AfterEdit(ByVal row As Long, ByVal col As Long)
With oGrid1
Select Case col
Case 5
    If oGrid1.TextMatrix(row, 5) >= oGrid1.TextMatrix(row, 3) And oGrid1.TextMatrix(row, 5) <= oGrid1.TextMatrix(row, 4) Then
        Text1(6).Text = oGrid1.TextMatrix(row, 5) 'iif( ogrid1.TextMatrix(Row, 5)=0,"",1)
    Else
        MsgBox "Entrian tidak sesuai dengan Level ", vbInformation
    End If
End Select
End With
End Sub

Private Sub ogrid1_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = oGrid1.Cols - 1 Then Exit Sub
If oGrid1.TextMatrix(row, oGrid1.Cols - 1) = "1" Then Exit Sub
GridModul.GridDetail_CellChanged row, col, oGrid1, istatus
End Sub

Private Sub oGrid1_Click()
'With ogrid1
'If .row = 0 Then Exit Sub
'Select Case .col
'Case 0
'    Text1(6) = ""
'    oResetGrid
'    Dim irow As Integer
'    For irow = 1 To .row
'        .TextMatrix(irow, 6) = "2"
'        .TextMatrix(irow, 7) = "1"
'    Next
'        .TextMatrix(.row, 6) = "1"
'        .TextMatrix(.row, 7) = "2"
'        If .TextMatrix(.row, .Cols - 1) = "0" Then
'            .TextMatrix(.row, .Cols - 1) = "2"
'        End If
'    .Select .row, 0
'    .Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
'    .EditCell
'    'Dim iRow As Integer
'Case 5
'    If .TextMatrix(.row, 0) = -1 Then
'        .Select .row, 5
'        .EditCell
'    End If
'End Select
'End With
End Sub

Private Sub Option4_Click(Index As Integer)
Select Case Index
Case 0
        skodegroup = oFindByQuery("select kodegroup from master_default_pelajaran where pelajaran='" & "1" & "'", parkir)
Case 1
        skodegroup = oFindByQuery("select kodegroup from master_default_pelajaran where pelajaran='" & "2" & "'", parkir)
Case 2
        skodegroup = oFindByQuery("select kodegroup from master_default_pelajaran where pelajaran='" & "3" & "'", parkir)
End Select


    ShowGrid1 skodegroup, Trim(Text1(5))
    
End Sub

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "mnUtama"
            Frame1(0).ZOrder   'Picture1(0).ZOrder
Case "mnDataSebelumnya"
            Frame1(1).ZOrder  'Picture1(1).ZOrder

End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "TabStrip1"
End Sub

Private Sub Text1_Change(Index As Integer)
Dim skriteriaku As String
skriteriaku = "where kdgroup='" & skodegroup & "' and kodelevel='" & ""
Select Case Index
Case 5
    Text1(5) = UCase(Text1(5))
    skriteriaku = "where kodegroup='" & skodegroup & "' and kodelevel='" & Trim(Text1(5)) & "'"
    lblgrouppelajaran.Caption = oFindByQuery("select namalevel from master_pelajaran_level " & skriteriaku, parkir)
    ShowGrid1 skodegroup, Trim(Text1(5))
    oNyariPosisi
Case 6
        oNyariPosisi
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000F

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
Select Case Index
Case 6
    If KeyCode = 13 Then
        oResetGrid
        oGrid1.TextMatrix(fceklevel(ToNumber(Text1(6))), 0) = 1
        oGrid1.TextMatrix(fceklevel(ToNumber(Text1(6))), 5) = ToNumber(Text1(6))
        oGrid1.Cell(flexcpBackColor, fceklevel(ToNumber(Text1(6))), 0, , oGrid1.Cols - 1) = vbGreen
        oGrid1.Refresh

    End If
End Select

End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).Text
End Sub
Public Sub ShowGrid1(skode1 As String, skode2 As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    sCekDataDetail1 = "0"
    sCekDataDetail1 = oFindByQuery("select count(*) from master_kelas_detail1 Where kodegroup='" & skode1 & "' and kodelevel='" & skode2 & "'", parkir)
    
    
    sQuery = "    SELECT  b.docentry, b.nokursus, b.tgltest, a.kodelevelno,a.kodelevel, a.kodegroup, "
    sQuery = sQuery & "   IFNULL(b.kodelevelnodetail,0) AS kodelevelnodetail, a.namalevelno,a.nolvlmulai,a.nolvlselesai,"
    sQuery = sQuery & "   IFNULL(b.waktupengerjaan,'') AS waktupengerjaan, "
    sQuery = sQuery & "   IFNULL(b.jawabanbenar,'') AS jawabanbenar, "
    sQuery = sQuery & "   IFNULL(b.titikpangkal,'') AS titikpangkal, "
    sQuery = sQuery & "   IFNULL(b.kelompok,'') AS kelompok, "
    sQuery = sQuery & "   IFNULL(b.aktif,'0') AS aktif, "
    sQuery = sQuery & "   IFNULL(b.status,'1') AS status,"
    sQuery = sQuery & "   IF(IFNULL(b.nokursus,'')='','1','0') AS keyentry    "
    sQuery = sQuery & "    "
    sQuery = sQuery & "   FROM "
    sQuery = sQuery & "   master_pelajaran_level_detail  a"
    sQuery = sQuery & "   LEFT JOIN master_kelas_detail1 b "
    sQuery = sQuery & "   ON a.kodegroup=b.kodegroup AND a.kodelevel=b.kodelevel AND a.kodelevelno=b.kodelevelno  AND b.docentry=" & sdocentry
    sQuery = sQuery & "   WHERE a.kodegroup='" & skode1 & "' and a.kodelevel='" & skode2 & "' order by nolvlmulai ,nolvlselesai"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodelevelno"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("namalevelno"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("nolvlmulai"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("nolvlselesai"))
                .TextMatrix(i, 5) = IIf(oRsDetail("kodelevelnodetail") = 0, "", oRsDetail("kodelevelnodetail"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("status"))
                .TextMatrix(i, .Cols - 3) = Trim(oRsDetail("aktif"))
                .TextMatrix(i, .Cols - 1) = RTrim(oRsDetail("keyentry"))
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                 
                 If (oRsDetail("aktif")) = "1" Then
                    .TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = vbGreen
                 End If
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

Public Function fceklevel(iCariPos As Integer) As Integer
Dim irow As Integer
Dim iPosRow As Integer
'Dim iCariPos As Integer
With oGrid1
    For irow = 1 To .Rows - 1
        If iCariPos >= ToNumber(.TextMatrix(irow, 3)) And iCariPos <= ToNumber(.TextMatrix(irow, 4)) Then
            fceklevel = irow
            irow = .Rows - 1
        End If
    Next
End With
End Function

Public Sub oResetGrid()
Dim irow As Integer
Dim iPosRow As Integer
'Dim iCariPos As Integer
With oGrid1
    For irow = 1 To .Rows - 1
        .TextMatrix(irow, 0) = 0
        .TextMatrix(irow, 5) = ""
        .TextMatrix(irow, 6) = "0"
        .TextMatrix(irow, 7) = "0"
        oGrid1.Cell(flexcpBackColor, irow, 0, , oGrid1.Cols - 1) = vbNormal
    Next
End With
End Sub

Public Sub oSaveDetail1()
Dim irow As Integer
sdocentry = oFindByQuery("select docentry from master_kelas where nokursus='" & Text1(0) & "'", parkir)
If skodelevellama <> skodelevel Then
    oCon.Execute "delete from master_kelas_detail1 where docentry='" & sdocentry & "'"
End If
With oGrid1
For irow = 1 To .Rows - 1
    ' set save detail1
    stgltest = stglmulai
    skodelevelno = .TextMatrix(irow, 1)
    skodelevelnodetail = ToNumber(.TextMatrix(irow, 5))
    saktif = .TextMatrix(irow, .Cols - 3)
    sstatus = .TextMatrix(irow, .Cols - 2)

    sUpdateDetail1 = "update  master_kelas_detail1   "
    sUpdateDetail1 = sUpdateDetail1 & "set "
    sUpdateDetail1 = sUpdateDetail1 & "nokursus='" & snokursus & "',"
    sUpdateDetail1 = sUpdateDetail1 & "tgltest='" & stgltest & "',"

    sUpdateDetail1 = sUpdateDetail1 & "kodelevelnodetail='" & skodelevelnodetail & "',"
    sUpdateDetail1 = sUpdateDetail1 & "waktupengerjaan='" & swaktupengerjaan & "',"
    sUpdateDetail1 = sUpdateDetail1 & "jawabanbenar='" & sjawabanbenar & "',"
    sUpdateDetail1 = sUpdateDetail1 & "titikpangkal='" & stitikpangkal & "',"
    sUpdateDetail1 = sUpdateDetail1 & "kelompok='" & skelompok & "',"
    sUpdateDetail1 = sUpdateDetail1 & "aktif='" & saktif & "',"
    sUpdateDetail1 = sUpdateDetail1 & "status='" & sstatus & "',"
    sUpdateDetail1 = sUpdateDetail1 & "audituser='" & saudituser & "',"
    sUpdateDetail1 = sUpdateDetail1 & "auditdate='" & sauditdate & "'"
    sUpdateDetail1 = sUpdateDetail1 & "where "
    sUpdateDetail1 = sUpdateDetail1 & "docentry='" & sdocentry & "' and "
    sUpdateDetail1 = sUpdateDetail1 & "kodelevelno='" & skodelevelno & "' and "
    sUpdateDetail1 = sUpdateDetail1 & "kodelevel='" & skodelevel & "' and "
    sUpdateDetail1 = sUpdateDetail1 & "kodegroup='" & skodegroup & "' "
    
    sInsertDetail1 = "insert into  master_kelas_detail1 "
    sInsertDetail1 = sInsertDetail1 & "("
    sInsertDetail1 = sInsertDetail1 & "docentry,nokursus,"
    sInsertDetail1 = sInsertDetail1 & "tgltest,"
    sInsertDetail1 = sInsertDetail1 & "kodelevelno,"
    sInsertDetail1 = sInsertDetail1 & "kodelevel,"
    sInsertDetail1 = sInsertDetail1 & "kodegroup,"
    sInsertDetail1 = sInsertDetail1 & "kodelevelnodetail,"
    sInsertDetail1 = sInsertDetail1 & "waktupengerjaan,"
    sInsertDetail1 = sInsertDetail1 & "jawabanbenar,"
    sInsertDetail1 = sInsertDetail1 & "titikpangkal,"
    sInsertDetail1 = sInsertDetail1 & "kelompok,"
    sInsertDetail1 = sInsertDetail1 & "aktif,"
    sInsertDetail1 = sInsertDetail1 & "status,"
    sInsertDetail1 = sInsertDetail1 & "audituser,"
    sInsertDetail1 = sInsertDetail1 & "auditdate"
    sInsertDetail1 = sInsertDetail1 & ")"
    sInsertDetail1 = sInsertDetail1 & " values "
    sInsertDetail1 = sInsertDetail1 & "('"
    sInsertDetail1 = sInsertDetail1 & sdocentry & "','"
    sInsertDetail1 = sInsertDetail1 & snokursus & "','"
    sInsertDetail1 = sInsertDetail1 & stgltest & "','"
    sInsertDetail1 = sInsertDetail1 & skodelevelno & "','"
    sInsertDetail1 = sInsertDetail1 & skodelevel & "','"
    sInsertDetail1 = sInsertDetail1 & skodegroup & "','"
    sInsertDetail1 = sInsertDetail1 & skodelevelnodetail & "','"
    sInsertDetail1 = sInsertDetail1 & swaktupengerjaan & "','"
    sInsertDetail1 = sInsertDetail1 & sjawabanbenar & "','"
    sInsertDetail1 = sInsertDetail1 & stitikpangkal & "','"
    sInsertDetail1 = sInsertDetail1 & skelompok & "','"
    sInsertDetail1 = sInsertDetail1 & saktif & "','"
    sInsertDetail1 = sInsertDetail1 & sstatus & "','"
    sInsertDetail1 = sInsertDetail1 & saudituser & "','"
    sInsertDetail1 = sInsertDetail1 & sauditdate & "'"
    sInsertDetail1 = sInsertDetail1 & ")"
    
    Select Case .TextMatrix(irow, .Cols - 1)
    Case "1"
        
        oCon.Execute sInsertDetail1
    Case "2"
        oCon.Execute sUpdateDetail1
    End Select
Next
End With


End Sub

Public Sub oSetToolBar(istatus As StatusForm, iStsDok As String)
        istatus = Normal
        MenuFrm.SetToolbar istatus
        If iStsDok = "1" Then
            MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = True
        Else
            MenuFrm.Toolbar1.Buttons(btm_Save).Enabled = False
        End If
End Sub

Public Sub oTampildrReff(sreffno As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select  *  from vtransaksi_pendaftaran where nopendaftaran ='" & sreffno & "' "
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showDatadrReff
    End If
    oCon.Close
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Tampil dari Reff"
End Sub

Public Sub showDatadrReff()
        Text1(1) = ToText(oRs("noidsiswa"))
        Text1(2) = ToText(oRs("nmlengkap"))
        Text1(3) = ToText(oRs("almtrumah1"))
        FlatDatePicker1(0).value = oRs("tglmasuk")
        Select Case ToText(oRs("penggolongan"))
        Case "1"
            Option3(0).value = True
        Case "2"
            Option3(1).value = True
        Case "3"
            Option3(2).value = True
        End Select
        Select Case ToText(oRs("pelajaran"))
        Case "1"
            Option4(0).value = True
        Case "2"
            Option4(1).value = True
        Case "3"
            Option4(2).value = True
        End Select
End Sub

Public Sub oSimpanMasterPembimbing()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "insert into master_pembimbing (namapembimbing) values ('" & snamapembimbing & "')"
    Set oRs = oCon.Execute(sQuery)

Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Tampil dari Reff"
End Sub

Public Sub oCreateKartuKelas(sdocentry As Integer)
On Error GoTo errhandler
Dim irow As Integer
Dim iRow2 As Integer

Dim slinenum As Integer
Dim slinenum2 As Integer
Dim sOk As Integer
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    
    With oGrid1
    slinenum = 1
    sOk = 1
    Dim sstatus2 As String
    Dim saktif2 As String
    For irow = 1 To .Rows - 1
        
        skodelevelno = .TextMatrix(irow, 1)
        saktif = .TextMatrix(irow, .Cols - 3)
        If sOk = 1 Then
            sstatus = .TextMatrix(irow, .Cols - 2)
        Else
            sstatus = "0"
        End If
        If sstatus = "2" Then
            sOk = 0
        Else
            sOk = 1
        End If
        
        sQuery = "insert into master_kartu_kelas_detail1"
        sQuery = sQuery & "(docentry,"
        sQuery = sQuery & "linenum,"
        sQuery = sQuery & "kodelevelno,"
        sQuery = sQuery & "kodelevel,"
        sQuery = sQuery & "kodegroup,"
        sQuery = sQuery & "aktif,"
        sQuery = sQuery & "status,"
        sQuery = sQuery & "audituser,"
        sQuery = sQuery & "auditdate"
        sQuery = sQuery & ")"
        sQuery = sQuery & " values "
        sQuery = sQuery & "("
        sQuery = sQuery & "'" & sdocentry & "',"
        sQuery = sQuery & "'" & slinenum & "',"
        sQuery = sQuery & "'" & skodelevelno & "',"
        sQuery = sQuery & "'" & skodelevel & "',"
        sQuery = sQuery & "'" & skodegroup & "',"
        sQuery = sQuery & "'" & saktif & "',"
        sQuery = sQuery & "'" & sstatus & "',"
        sQuery = sQuery & "'" & MenuFrm.sUserID & "',"
        sQuery = sQuery & "'" & Format(Now(), "YYYY/MM/DD") & "'"
        sQuery = sQuery & ")"
        
        Set oRs = oCon.Execute(sQuery)
        Dim sbaselinenum As Integer
        sbaselinenum = slinenum
        slinenum2 = 1
        For iRow2 = ToNumber(.TextMatrix(irow, 3)) To ToNumber(.TextMatrix(irow, 4))
            If ToNumber(ToNumber(Text1(6).Text)) = iRow2 Then
               swaktupengerjaan = Text1(7).Text
               sjawabanbenar = Text1(8).Text
               stitikpangkal = Text1(9).Text
               skelompok = Text1(10).Text
               saktif2 = "1"
               sstatus2 = "2"
            Else
            If ToNumber(ToNumber(Text1(6).Text)) > iRow2 Then
               swaktupengerjaan = ""
               sjawabanbenar = ""
               stitikpangkal = ""
               skelompok = ""
               saktif2 = "2"
               sstatus2 = "1"
            Else
                swaktupengerjaan = ""
               sjawabanbenar = ""
               stitikpangkal = ""
               skelompok = ""
               saktif2 = "0"
               sstatus2 = "0"
            End If
            End If
            skodelevelnodetail = iRow2
            sQuery = "insert into  master_kartu_kelas_detail2"
            sQuery = sQuery & "("
            sQuery = sQuery & "docentry,"
            sQuery = sQuery & "baselinenum,"
            sQuery = sQuery & "linenum,"
            sQuery = sQuery & "kodelevelnodetail,waktupengerjaan,jawabanbenar,"
            sQuery = sQuery & "titikpangkal,kelompok,aktif,"
            sQuery = sQuery & "status,"
            sQuery = sQuery & "audituser,"
            sQuery = sQuery & "auditdate"
            sQuery = sQuery & ")"
            sQuery = sQuery & " values "
            sQuery = sQuery & "('"
            sQuery = sQuery & sdocentry & "','"
            sQuery = sQuery & sbaselinenum & "','"
            sQuery = sQuery & slinenum2 & "','"
            sQuery = sQuery & iRow2 & "','" & swaktupengerjaan & "','" & sjawabanbenar & "','"
            sQuery = sQuery & stitikpangkal & "','"
            sQuery = sQuery & skelompok & "','"
            sQuery = sQuery & saktif2 & "','"
            sQuery = sQuery & sstatus2 & "','"
            sQuery = sQuery & saudituser & "','"
            sQuery = sQuery & sauditdate & "'"
            sQuery = sQuery & ")"
            Set oRs = oCon.Execute(sQuery)
            slinenum2 = slinenum2 + 1
        Next
        slinenum = slinenum + 1
    Next
    End With
    oCreateKartuBayar
    sQuery = "update master_kelas set dokstatus=0 where docentry='" & sdocentry & "'"
    Set oRs = oCon.Execute(sQuery)
    oCon.Close
    FindData snokursus
    Command1.Enabled = False
    oSetToolBar Normal, sdokstatus
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "oCreateKartuKelas"
End Sub

Public Sub oCreateKartuBayar()
Dim smop As Integer
Dim syop As Integer
Dim irow As Integer

smop = Month(FlatDatePicker1(0).value)
syop = Year(FlatDatePicker1(0).value)

If (Year(FlatDatePicker1(0).value) * 100 + Month(FlatDatePicker1(0).value)) <= (Year(Now()) * 100 + Month(Now())) Then
    smop = Month(Now())
    syop = Year(Now())
End If

For irow = 1 To 12
        If smop > irow Then
           sstsbayar = "0"
           sjnsbayar = "1"
        Else
            sstsbayar = "1"
            sjnsbayar = "0"
        End If
        sQuery = "insert into master_kartu_bayar"
        sQuery = sQuery & "("
        sQuery = sQuery & "docentry,"
        sQuery = sQuery & "nokursus,"
        sQuery = sQuery & "yop,"
        sQuery = sQuery & "mop,"
        sQuery = sQuery & "basedocentry,"
        sQuery = sQuery & "nokwitansi,"
        sQuery = sQuery & "jnsbayar,"
        sQuery = sQuery & "stsbayar,"
        sQuery = sQuery & "nilaibayar,"
        sQuery = sQuery & "tglbayar,"
        sQuery = sQuery & "audituser,"
        sQuery = sQuery & "auditdate"
        sQuery = sQuery & ")"
        sQuery = sQuery & " values "
        sQuery = sQuery & "('"
        sQuery = sQuery & sdocentry & "','"
        sQuery = sQuery & snokursus & "','"
        sQuery = sQuery & syop & "','"
        sQuery = sQuery & irow & "','"
        sQuery = sQuery & 0 & "','"
        sQuery = sQuery & "" & "','"
        sQuery = sQuery & sjnsbayar & "','"
        sQuery = sQuery & sstsbayar & "','"
        sQuery = sQuery & 0 & "','"
        sQuery = sQuery & "0000-00-00" & "','"
        sQuery = sQuery & saudituser & "','"
        sQuery = sQuery & sauditdate & "'"
        sQuery = sQuery & ")"
        Set oRs = oCon.Execute(sQuery)
Next

End Sub

Public Sub oNyariPosisi()
        If fceklevel(ToNumber(Text1(6))) = 0 Then Exit Sub
        oResetGrid
        
        With oGrid1
        Dim irow As Integer
        For irow = 1 To fceklevel(ToNumber(Text1(6)))
        .TextMatrix(irow, 6) = "2"
        .TextMatrix(irow, 7) = "1"
        Next

        .TextMatrix(fceklevel(ToNumber(Text1(6))), .Cols - 2) = "2"
        .TextMatrix(fceklevel(ToNumber(Text1(6))), .Cols - 3) = "1"
        oGrid1.TextMatrix(fceklevel(ToNumber(Text1(6))), 0) = 1
        oGrid1.TextMatrix(fceklevel(ToNumber(Text1(6))), 5) = ToNumber(Text1(6))
        
        oGrid1.Cell(flexcpBackColor, fceklevel(ToNumber(Text1(6))), 0, , oGrid1.Cols - 1) = vbGreen
        oGrid1.Refresh
        End With
End Sub
Public Sub Execution()
On Error GoTo errhandler
Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\KelasFrm.Rpt"

Dim sKriteria As String

sQuery = "SELECT * from vmaster_kelas_rpt vmaster_kelas_rpt1 where nokursus='" & Text1(0) & "'"

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

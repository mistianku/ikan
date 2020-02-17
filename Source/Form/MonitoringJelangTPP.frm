VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MonitoringJelangTPPFrm 
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
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
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
      Left            =   13200
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cari Berdasarkan"
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   40
      Top             =   720
      Width           =   11775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5160
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   200
         Width           =   6495
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Nama Siswa"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   43
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No.ID.Siswa"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   42
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "No.Kursus"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kelas"
      Height          =   2775
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid3 
         Height          =   2415
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   4260
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
         BackColorSel    =   65408
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringJelangTPP.frx":0000
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
   Begin Crystal.CrystalReport cr1 
      Left            =   12240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   8400
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Level"
            Key             =   "mnLevel"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Level Detail"
            Key             =   "mnLevelDetail"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Materi"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   4080
      Width           =   11775
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   31
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MonitoringJelangTPP.frx":00F9
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
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label Label1 
         Caption         =   " Group Level"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   735
      Index           =   1
      Left            =   12120
      TabIndex        =   0
      Top             =   5760
      Visible         =   0   'False
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
         Left            =   5640
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   6015
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   8280
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         Left            =   5040
         TabIndex        =   29
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MonitoringJelangTPP.frx":0115
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
         Index           =   2
         Left            =   8400
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   2775
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
         Left            =   2280
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2775
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   11160
         TabIndex        =   30
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "MonitoringJelangTPP.frx":0131
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
         Caption         =   " Tanggal Mulai"
         Height          =   315
         Index           =   3
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   " No.Kursus"
         Height          =   315
         Index           =   2
         Left            =   6240
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   " Nama"
         Height          =   315
         Index           =   1
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   " No.ID Siswa"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000A&
      Caption         =   "Jenis Tes"
      Height          =   615
      Left            =   9240
      TabIndex        =   47
      Top             =   4080
      Width           =   2655
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Reguler"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Ulang"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   48
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kelas"
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   11775
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inggris EFL"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   39
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Inggris EE"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Matematika"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Level Info"
      Enabled         =   0   'False
      Height          =   3615
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11535
         _cx             =   20346
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringJelangTPP.frx":014D
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
      Caption         =   "Level Detail No Info"
      Enabled         =   0   'False
      Height          =   3615
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid2 
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   4895
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringJelangTPP.frx":02B8
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   9
         Left            =   9720
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   7
         Left            =   5880
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   6
         Left            =   3960
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Index           =   8
         Left            =   7800
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1900
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   33
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label lblgede 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   11535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Titik Pangkal"
         Height          =   315
         Index           =   11
         Left            =   7800
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Kelompok"
         Height          =   315
         Index           =   10
         Left            =   9720
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Titik Pangkal"
         Height          =   315
         Index           =   9
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Jawaban Benar"
         Height          =   315
         Index           =   8
         Left            =   5880
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Waktu Pengerjaan"
         Height          =   315
         Index           =   7
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Level"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1900
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   50
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   " Kelas"
      Height          =   315
      Index           =   4
      Left            =   12000
      TabIndex        =   46
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "MonitoringJelangTPPFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sLebarQ As Integer
Dim sLebarLbl As Integer
Dim sdocentry As Integer
Dim slinenum As Integer
Dim skodelevelno As String
Dim skodelevel As String
Dim skodegroup As String
Dim saktif As String
Dim sstatus As String

Dim saudituser As String
Dim sauditdate As String

Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Dim sKriteria As String
Dim sSortbyQ As String

Dim snoidsiswa As String
Dim snmlengkap As String
Dim snokursus As String
Dim stglmulai As String
Dim sstskelas As String
Dim spelajaran As String
Dim sketerangan As String


Dim sbaselinenum As Integer

Dim skodelevelnodetail As Integer
Dim stgltest As String
Dim swaktupengerjaan As String
Dim sjawabanbenar As String
Dim stitikpangkal As String
Dim skelompok As String

Dim sAktifSeb As String
Dim sStatusSeb As String

Dim sKeyAwal As Integer

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from vmaster_kartu_materi_kelas  where stskelas='1' and noidsiswa='" & sKodeUserAkses & "' order by noidsiswa,nokursus asc"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
        MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
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
    sQuery = "Select *  from vmaster_kartu_materi_kelas where stskelas='1' order by noidsiswa,nokursus asc limit 1"
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
    sQuery = "Select  *  from vmaster_kartu_materi_kelas where stskelas='1' and noidsiswa >'" & Text1(0).Text & "' order by noidsiswa asc ,nokursus asc limit 1"
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
    sQuery = "Select  *  from vmaster_kartu_materi_kelas where stskelas='1' and  noidsiswa<'" & Text1(0).Text & "' order by noidsiswa desc ,nokursus desc limit 1"
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
    sQuery = "Select *  from vmaster_kartu_materi_kelas order by noidsiswa desc limit 1 "
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
'Dim ires As Integer
'    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
'    If ires = 6 Then
'        If DoSaveData Then
'             ShowGrid1 sdocentry, skodelevel
'             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
'             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
'             MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
'        End If
'    End If
End Sub
Public Sub DeleteData()
'    Dim ires As Integer
'    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
'    If ires = 6 Then
'        If DoDeleteData Then
'             MsgBox "Data Sudah Terhapus", , "Hapus Data"
'             MovePrevious
'        End If
'    End If
'    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
'    MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
'    If setData Then
'        If oCon.State = 1 Then oCon.Close
'         oCon.Open MainModule.Conectionku(DBaseConection.Parkir)
'        If istatus = StatusForm.DataBaru Then
'        sQuery = sInsert
'        Else
'        sQuery = sUpdate
'        End If
'        oCon.Execute sQuery
'        oCon.Close
'        DoSaveData = True
'        istatus = Normal
'        Exit Function
'    End If
oSaveKartuDetail1
oSaveKartuDetail2
DoSaveData = True
Exit Function
errhandler:
MainModule.ShowMessage Err.Description, "savedata"
End Function
Private Function DoDeleteData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "Delete from vmaster_kartu_materi_kelas where noidsiswa='" & snoidsiswa & "'"
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
'    KodeUserAksesTemp = Text1(0)
'    istatus = StatusForm.DataBaru
'    'cleardata
'    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
'    Text1(0).Locked = False
'    Text1(0).SetFocus
'    Text1(0).TabIndex = 0
'    Text1(1).TabIndex = 1
'    Text1(4) = ""
'    Text1(4).SetFocus
End Sub
Public Sub Undo()
'    istatus = Normal
'    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
'    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    snoidsiswa = Text1(0).Text
    sagama = Text1(1).Text
     
    sUpdate = "update master_agama set "
    sUpdate = sUpdate & "agama='" & sagama & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "'"
    
    sInsert = "insert into master_agama ("
    sInsert = sInsert & "noidsiswa,agama ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & snoidsiswa & "',"
    sInsert = sInsert & "'" & sagama & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    
'    oBrowse.ShowFinder BrowsSiswa, "stssiswa='1'" '" EXISTS(SELECT * FROM master_kelas WHERE vmaster_siswa.noidsiswa=noidsiswa)"
'    If Not oBrowse.YangDipilih = "" Then
'        Text1(0) = oBrowse.YangDipilih
'        Text1(0) = oBrowse.Keterangan
'        ShowGridKelas oBrowse.YangDipilih
'    End If
'    If Not oBrowse.YangDipilih = "" Then
'        NewData
'        FindData oBrowse.YangDipilih
'    End If
Case 1
    oBrowse.ShowFinder BrowsKartuKelas, "noidsiswa='" & Text1(0) & "'"
     If Not oBrowse.YangDipilih = "" Then
        Text1(2).Text = oBrowse.YangDipilih
        FlatDatePicker1(0).value = ToDate(oBrowse.Keterangan)
        Text1(3) = oFindByQuery("select keterangan from vmaster_kartu_materi_kelas where nokursus='" & Text1(2) & "'", parkir)
        Text1(4) = oFindByQuery("select kodelevel from vmaster_kartu_materi_kelas where nokursus='" & Text1(2) & "'", parkir)
        sdocentry = oFindByQuery("select docentry from vmaster_kartu_materi_kelas where nokursus='" & Text1(2) & "'", parkir)
        skodegroup = oFindByQuery("select kodegroup from vmaster_kartu_materi_kelas where nokursus='" & Text1(2) & "'", parkir)
        skodelevel = Text1(4)
        Text1(4).SetFocus
        ShowGrid1 sdocentry, skodelevel
    End If
Case 2
      sKeyAwal = 1
      oBrowse.ShowFinder BrowsPelajaranLevel, "kodegroup='" & skodegroup & "'"
      If Not oBrowse.YangDipilih = "" Then
            Text1(4) = oBrowse.YangDipilih
            Text1(4).SetFocus
      End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub FlatDatePicker1_DatePick(Index As Integer, CurrentDate As Date)
Select Case Index
Case 0
ogrid2.TextMatrix(ogrid2.row, 2) = FlatDatePicker1(1).value
Case 1
End Select
End Sub

Private Sub FlatDatePicker1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Select Case Index
Case 0

Case 1
ogrid2.TextMatrix(ogrid2.row, 2) = FlatDatePicker1(1).value
End Select
End Sub

Private Sub FlatDatePicker1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0

Case 1
    ogrid2.TextMatrix(ogrid2.row, 2) = FlatDatePicker1(1).value
End Select
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Monitoring Jelang TPP "
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas

MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width
End Sub

Private Sub Form_Load()
sLebarQ = Frame1(0).Width
sLebarLbl = Label4.Width
sKeyAwal = 0
oFormatFrameBackground Frame4
oFormatOption 3, Me

cleardata
istatus = Normal
MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
Frame4.Visible = False
spelajaran = 1
If Option3(0).value = True Then
    sKriteria = " nokursus like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "nokursus"
End If
If Option3(1).value = True Then
    sKriteria = " noidsiswa like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "noidsiswa"
End If
If Option3(2).value = True Then
    sKriteria = " nmlengkap like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "nmlengkap"
End If
MoveFirst
Frame1(2).ZOrder
Label4.Caption = oFindByQuery("select namalevel from master_pelajaran_level where kodegroup='" & skodegroup & "' and kodelevel='" & skodelevel & "'", parkir)
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Label3.Caption = ToNumber(oRs("docentry"))
    sdocentry = ToNumber(oRs("docentry"))
    Text1(0).Text = oRs("noidsiswa")
    KodeUserAksesTemp = oRs("noidsiswa")
    Text1(0).Locked = True
    Text1(1).Text = oRs("nmlengkap")
    Text1(2).Text = oRs("nokursus")
    Text1(3).Text = oRs("keterangan")
    FlatDatePicker1(0).value = oRs("tglmulai")
    snoidsiswa = oRs("noidsiswa")
    skodegroup = oRs("kodegroup")
    skodelevel = oRs("kodelevel")
    Text1(4).Text = skodelevel
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    ShowGridKelas spelajaran, sKriteria, sSortbyQ
    ShowGrid1 sdocentry, skodelevel
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

Private Sub Label1_Click(Index As Integer)
2250
End Sub

Private Sub Label3_Change()
ClearGridDetail oGrid1
ClearGridDetail ogrid2
oGetMateriku Label3.Caption
sdocentry = Label3.Caption
skodelevel = Text1(4)
    If Not ToNumber(oFindByQuery("select count(*) from master_kartu_kelas_detail1 where docentry=" & ToNumber(Label3.Caption) & " and kodegroup='" & skodegroup & "' and kodelevel='" & skodelevel & "'", parkir)) = 0 Then
        ShowGrid1 ToNumber(Label3.Caption), skodelevel
    Else
        If sKeyAwal = 0 Then Exit Sub
        Dim slinenumstart As Integer
        GridModul.ClearGridDetail oGrid1
        If MsgBox("Data Belum Ada !!, Akan Ditambahkan ", vbYesNo) = vbNo Then Exit Sub
        slinenumstart = ToNumber(oFindByQuery("select max(linenum) from master_kartu_kelas_detail1 where docentry=" & sdocentry, parkir))
        slinenumstart = slinenumstart + 1
        ShowGrid1_newdata sdocentry, skodelevel, slinenumstart
        sKeyAwal = 0
    End If
End Sub

Private Sub ogrid1_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = oGrid1.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, oGrid1, istatus
If oGrid1.TextMatrix(row, 5) = "1" Then
    ShowGrid2 sdocentry, oGrid1.TextMatrix(oGrid1.row, oGrid1.Cols - 2)
End If
If oGrid1.TextMatrix(row, 0) = -1 And Not row = 0 Then
    lblgede = "Level (" & oGrid1.TextMatrix(oGrid1.row, 1) & ") " & oGrid1.TextMatrix(oGrid1.row, 2)
End If
End Sub

Private Sub oGrid1_Click()
With oGrid1
Dim irow As Integer
For irow = 1 To .Rows - 1
    
    .Cell(flexcpBackColor, irow, 0, , .Cols - 1) = vbNormal   '&HC0C0FF
    If irow = .row And .TextMatrix(irow, 0) = -1 Then
        .TextMatrix(irow, 0) = -1
        .TextMatrix(irow, 7) = "1"
    Else
        .TextMatrix(irow, 0) = 0
        If irow > .row Then
            .TextMatrix(irow, 7) = "0"
            .TextMatrix(irow, 8) = "0"
            .TextMatrix(irow, 0) = 0
        End If
    End If
    Me.Caption = .row
    If .TextMatrix(irow, 6) = "2" Then
        .Cell(flexcpBackColor, irow, 0, , .Cols - 1) = &H8000000A
    Else
        .Cell(flexcpBackColor, irow, 0, , .Cols - 1) = vbNormal   '&HC0C0FF
    End If
Next
If .row = 0 Then Exit Sub
Select Case .col
Case 0
    'If .TextMatrix(.Row, .Cols - 5) = "0" Then
        .Select .row, .col
        .EditCell
        .TextMatrix(.row, .Cols - 5) = 1
        .TextMatrix(.row, .Cols - 4) = 2
        If Not .row - 1 = 0 Then
            .Cell(flexcpBackColor, .row - 1, 0, , .Cols - 1) = vbNormal '&HC0C0FF
            .TextMatrix(.row - 1, .Cols - 5) = 2
            .TextMatrix(.row - 1, .Cols - 4) = 1
        End If
        
        If .TextMatrix(.row, 0) = -1 Then
            '.Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
            If Not ToNumber(oFindByQuery("select count(*) from master_kartu_kelas_detail2 where docentry=" & sdocentry & " and baselinenum=" & .TextMatrix(.row, .Cols - 2), parkir)) = 0 Then
                ShowGrid2 sdocentry, .TextMatrix(.row, .Cols - 2)
            Else
                Dim sawal As Integer
                Dim sAkhir As Integer
                sawal = ToNumber(oFindByQuery("select nolvlmulai from master_pelajaran_level_detail where kodegroup='" & .TextMatrix(.row, 4) & "' and kodelevel='" & .TextMatrix(.row, 3) & "' and kodelevelno='" & .TextMatrix(.row, 1) & "'", parkir))
                sAkhir = ToNumber(oFindByQuery("select nolvlselesai from master_pelajaran_level_detail where kodegroup='" & .TextMatrix(.row, 4) & "' and kodelevel='" & .TextMatrix(.row, 3) & "' and kodelevelno='" & .TextMatrix(.row, 1) & "'", parkir))
                ShowGrid2_newdata sdocentry, .TextMatrix(.row, .Cols - 2), sawal, sAkhir
            End If
'            ogrid2.Cell(flexcpBackColor, .row, 0, , ogrid2.Cols - 1) = vbGreen
'            ogrid2.TextMatrix(.row, 0) = -1
'            ogrid2.TextMatrix(.row, ogrid2.Cols - 6) = 1
'            ogrid2.TextMatrix(.row, ogrid2.Cols - 5) = 2
            
            oGrid1.Cell(flexcpBackColor, .row, 0, , oGrid1.Cols - 1) = vbGreen
            oGrid1.TextMatrix(.row, 0) = -1
            oGrid1.TextMatrix(.row, oGrid1.Cols - 6) = 1
            oGrid1.TextMatrix(.row, oGrid1.Cols - 5) = 2
            
            .Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
        End If
End Select
End With
End Sub

Private Sub ogrid2_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = ogrid2.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, ogrid2, istatus
End Sub

Private Sub oGrid2_Click()
Dim irow As Integer
With ogrid2

For irow = 1 To .Rows - 1
    If .row = irow Then
        .TextMatrix(irow, 0) = -1
    Else
        .TextMatrix(irow, 0) = 0
    End If
    '.TextMatrix(iRow, 0) = 0
    If .TextMatrix(irow, 0) = -1 Then
        .TextMatrix(irow, .Cols - 6) = 2
        .TextMatrix(irow, .Cols - 5) = 1
    Else
        .TextMatrix(irow, .Cols - 6) = 0
        .TextMatrix(irow, .Cols - 5) = 0
    End If
    
    'If .TextMatrix(irow, .Cols - 7) = 2 And .TextMatrix(irow, .Cols - 8) = 1 Then
    If ToNumber(.TextMatrix(irow, .Cols - 8)) = 1 Then
        .TextMatrix(irow, .Cols - 6) = 2
        .TextMatrix(irow, .Cols - 5) = .TextMatrix(irow, .Cols - 7)
    Else
        .TextMatrix(irow, .Cols - 6) = .TextMatrix(irow, .Cols - 8)
        .TextMatrix(irow, .Cols - 5) = .TextMatrix(irow, .Cols - 7)
    End If
    .Cell(flexcpBackColor, irow, 0, , .Cols - 1) = vbNormal
    
Next
    Select Case .col
    Case 0
         .EditCell
        .Select .row, 0
       .TextMatrix(.row, 0) = -1
        If .TextMatrix(.row, 0) = -1 Then
            .TextMatrix(.row, 0) = -1
            .TextMatrix(.row, .Cols - 6) = 1
            .TextMatrix(.row, .Cols - 5) = 2
        Else
            .TextMatrix(.row, .Cols - 6) = .TextMatrix(.row, .Cols - 8)
            .TextMatrix(.row, .Cols - 5) = .TextMatrix(.row, .Cols - 7)
        End If
        If Option1(0).value = True Then
            .TextMatrix(.row, .Cols - 5) = 3
        Else
            .TextMatrix(.row, .Cols - 5) = 4
        End If
        
        .Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
         Text1(5) = .TextMatrix(.row, 1)
         Text1(6) = ogrid2.TextMatrix(ogrid2.row, 3)
         Text1(7) = ogrid2.TextMatrix(ogrid2.row, 4)
         Text1(8) = ogrid2.TextMatrix(ogrid2.row, 5)
         Text1(9) = ogrid2.TextMatrix(ogrid2.row, 6)
         

         Text1(6).Enabled = True
         Text1(7).Enabled = True
         Text1(8).Enabled = True
         Text1(9).Enabled = True
    End Select
End With
End Sub

Private Sub ogrid2_DblClick()
With ogrid2
If Not .TextMatrix(.row, 0) = -1 Then Exit Sub
EntryTestEvaluasiFrm.Show
End With
End Sub

Private Sub oGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

End Select
End Sub

Private Sub ogrid3_Click()
GridModul.ClearGridDetail ogrid2
With ogrid3
If .row = 0 Then Exit Sub
Label3.Caption = .TextMatrix(.row, 0)
End With
End Sub

Private Sub Option2_Click(Index As Integer)
cleardata
Label4 = ""
lblgede = ""
GridModul.ClearGridDetail oGrid1
GridModul.ClearGridDetail ogrid2
If Option2(0).value = True Then
    spelajaran = "1"
End If
If Option2(1).value = True Then
    spelajaran = "2"
End If
If Option2(2).value = True Then
    spelajaran = "3"
End If
ShowGridKelas spelajaran, sKriteria, sSortbyQ
End Sub

Private Sub Option3_Click(Index As Integer)
If Option3(0).value = True Then
    sKriteria = " nokursus like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "nokursus"
End If
If Option3(1).value = True Then
    sKriteria = " noidsiswa like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "noidsiswa"
End If
If Option3(2).value = True Then
    sKriteria = " nmlengkap like '%" & Trim(Text1(10)) & "%'"
    sSortbyQ = "nmlengkap"
End If
ShowGridKelas spelajaran, sKriteria, sSortbyQ
End Sub

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "mnLevel"
            Frame1(2).ZOrder   'Picture1(0).ZOrder
            Frame4.Visible = False
            MenuFrm.LblPesanku = "Pilih Level Detail dan Double Click dibaris Yang Akan di Entry Hasil Test !!!"
            Frame1(0).Width = sLebarQ
            Label4.Width = sLebarLbl
Case "mnLevelDetail"
            Frame1(3).ZOrder  'Picture1(1).ZOrder
            Frame4.Visible = True
            MenuFrm.LblPesanku = "Pilih Level Detail dan Double Click dibaris Yang Akan di Entry Hasil Test !!!"
            Label4.Width = sLebarLbl - Frame4.Width
            Frame1(0).Width = sLebarQ - Frame4.Width
End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "Tab Level"
End Sub

Private Sub Text1_Change(Index As Integer)
Dim sKey As Integer
Select Case Index
Case 2
            If Not Text1(2) = "" Then
                'FlatDatePicker1.CurrentDate = ToDate(oFindByQuery("select tglmulai from vmaster_kartu_materi_kelas where nokursus='" & Text1(2) & "'", Parkir))
                Text1(3) = oFindByQuery("select keterangan from vmaster_kartu_materi_kelas where nokursus='" & ToText(Text1(2)) & "'", parkir)
                Text1(4) = oFindByQuery("select kodelevel from vmaster_kartu_materi_kelas where nokursus='" & ToText(Text1(2)) & "'", parkir)
                sdocentry = ToNumber(oFindByQuery("select docentry from vmaster_kartu_materi_kelas where nokursus='" & ToText(Text1(2)) & "'", parkir))
                skodegroup = oFindByQuery("select kodegroup from vmaster_kartu_materi_kelas where nokursus='" & ToText(Text1(2)) & "'", parkir)
                skodelevel = ToText(Text1(4))
                'Text1(4).SetFocus
                ShowGrid1 sdocentry, skodelevel
            End If
Case 4
    skodelevel = ToText(Text1(4))
    If Not ToNumber(oFindByQuery("select count(*) from master_kartu_kelas_detail1 where docentry=" & ToNumber(Label3.Caption) & " and kodegroup='" & skodegroup & "' and kodelevel='" & skodelevel & "'", parkir)) = 0 Then
        ShowGrid1 ToNumber(Label3.Caption), skodelevel
    Else
        If sKeyAwal = 0 Then Exit Sub
        Dim slinenumstart As Integer
        GridModul.ClearGridDetail oGrid1
        If MsgBox("Data Belum Ada !!, Akan Ditambahkan ", vbYesNo) = vbNo Then Exit Sub
        slinenumstart = ToNumber(oFindByQuery("select max(linenum) from master_kartu_kelas_detail1 where docentry=" & sdocentry, parkir))
        slinenumstart = slinenumstart + 1
        ShowGrid1_newdata sdocentry, skodelevel, slinenumstart
        sKeyAwal = 0
    End If
    
Case 6
    If ogrid2.row = 0 Then Exit Sub
    ogrid2.TextMatrix(ogrid2.row, 3) = ToText(Text1(Index))
Case 7
    If ogrid2.row = 0 Then Exit Sub
    ogrid2.TextMatrix(ogrid2.row, 4) = ToText(Text1(Index))
Case 8
    If ogrid2.row = 0 Then Exit Sub
    ogrid2.TextMatrix(ogrid2.row, 5) = ToText(Text1(Index))
Case 9
    If ogrid2.row = 0 Then Exit Sub
    ogrid2.TextMatrix(ogrid2.row, 6) = ToText(Text1(Index))
Case 10

        If Option3(0).value = True Then
            sKriteria = " nokursus like '%" & ToText(Trim(Text1(10))) & "%'"
            sSortbyQ = "nokursus"
        End If
        If Option3(1).value = True Then
            sKriteria = " noidsiswa like '%" & ToText(Trim(Text1(10))) & "%'"
            sSortbyQ = "noidsiswa"
        End If
        If Option3(2).value = True Then
            sKriteria = " nmlengkap like '%" & ToText(Trim(Text1(10))) & "%'"
            sSortbyQ = "nmlengkap"
        End If
        ShowGridKelas spelajaran, sKriteria, sSortbyQ

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
'If Index = 0 Then FindData Text1(0).Text
End Sub

Public Sub ShowGrid1(keyDocentry As Integer, keyLevel As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim skeyBaseline As Integer
      
    sQuery = "select "
    sQuery = sQuery & "a.docentry,"
    sQuery = sQuery & "a.linenum,"
    sQuery = sQuery & "a.kodelevelno,b.namalevelno,"
    sQuery = sQuery & "a.kodelevel,"
    sQuery = sQuery & "a.kodegroup,"
    sQuery = sQuery & "a.aktif,"
    sQuery = sQuery & "a.status "
    sQuery = sQuery & " from "
    sQuery = sQuery & " master_kartu_kelas_detail1  as a"
    sQuery = sQuery & " inner join master_pelajaran_level_detail as b "
    sQuery = sQuery & " on "
    sQuery = sQuery & " a.kodelevelno=b.kodelevelno and "
    sQuery = sQuery & " a.kodelevel=b.kodelevel and "
    sQuery = sQuery & " a.kodegroup=b.kodegroup"
    sQuery = sQuery & "   WHERE a.aktif=1 and a.docentry='" & keyDocentry & "' and a.kodelevel='" & keyLevel & "' order by a.linenum "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)


    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodelevelno"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("namalevelno"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("kodelevel"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("kodegroup"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("aktif"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("status"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("aktif"))
                .TextMatrix(i, 8) = RTrim(oRsDetail("status"))
                .TextMatrix(i, .Cols - 3) = Trim(oRsDetail("docentry"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("linenum"))
                .TextMatrix(i, .Cols - 1) = 0
                
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
'                 If (oRsDetail("aktif")) = "2" Then
'                    .TextMatrix(i, 0) = -1
'                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = &HC0C0FF
'                 End If
                 If .TextMatrix(i, 7) = "1" Then
                    skeyBaseline = .TextMatrix(i, .Cols - 2)
                    .TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = vbGreen
                    ShowGrid2 .TextMatrix(i, .Cols - 3), .TextMatrix(i, .Cols - 2)
                    'ShowGrid2 .TextMatrix(i, .Cols - 3), .TextMatrix(i, .Cols - 2)
                    lblgede = "Level (" & .TextMatrix(i, 1) & ") " & .TextMatrix(i, 2)
                 End If
                 
                 If (oRsDetail("aktif")) = "2" Then
                    '.TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = &H8000000A
                   
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

Public Sub ShowGrid2(keyDocentry As Integer, keylinenum As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim iPosAktif As Integer
      
    sQuery = "select * from master_kartu_kelas_detail2 where  docentry=" & keyDocentry & " And baselinenum='" & keylinenum & "'"
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid2

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid2
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            iPosAktif = 1
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodelevelnodetail"))
                .TextMatrix(i, 2) = Format(oRsDetail("tgltest"), "MM/DD/YYYY")
                .TextMatrix(i, 3) = RTrim(oRsDetail("waktupengerjaan"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("jawabanbenar"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("titikpangkal"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("kelompok"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("aktif"))
                .TextMatrix(i, 8) = RTrim(oRsDetail("STATUS"))
                .TextMatrix(i, .Cols - 4) = Trim(oRsDetail("docentry"))
                .TextMatrix(i, .Cols - 3) = RTrim(oRsDetail("baselinenum"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("linenum"))
                .TextMatrix(i, .Cols - 1) = 0
                
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                 If (oRsDetail("aktif")) = "1" Then
                    .TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = &HC0C0FF
                 End If
                 If (oRsDetail("aktif")) = "1" Then
                    '.TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = vbGreen
                    iPosAktif = i
                 End If
                 If (oRsDetail("aktif")) = "2" Then
                    '.TextMatrix(i, 0) = -1
                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = &H8000000A
                    iPosAktif = i
                 End If
                 
                oRsDetail.MoveNext
            Loop
            '.Select iPosAktif, 0
            If .TextMatrix(iPosAktif, 0) = -1 Then
                Text1(6).Enabled = True
                Text1(7).Enabled = True
                Text1(8).Enabled = True
                Text1(9).Enabled = True
            Else
                Text1(6).Enabled = False
                Text1(7).Enabled = False
                Text1(8).Enabled = False
                Text1(9).Enabled = False
            End If
            Text1(5) = ogrid2.TextMatrix(iPosAktif, 1)
            Text1(6) = ogrid2.TextMatrix(iPosAktif, 2)
            Text1(7) = ogrid2.TextMatrix(iPosAktif, 3)
            Text1(8) = ogrid2.TextMatrix(iPosAktif, 4)
            Text1(9) = ogrid2.TextMatrix(iPosAktif, 5)
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub


Public Sub ShowGrid1_newdata(keyDocentry As Integer, keyLevel As String, keylinenum As Integer)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "select * from master_pelajaran_level_detail where kodegroup='" & skodegroup & "' and "
    sQuery = sQuery & " kodelevel='" & keyLevel & "'"
    sQuery = sQuery & " order by nolvlmulai,nolvlselesai"
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)


    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        '.ColHidden(.Cols - 1) = True
        Dim iLinenum As Integer
        iLinenum = keylinenum
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodelevelno"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("namalevelno"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("kodelevel"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("kodegroup"))
                .TextMatrix(i, 5) = 0
                .TextMatrix(i, 6) = 0
                .TextMatrix(.Rows - 1, .Cols - 5) = 0
                .TextMatrix(.Rows - 1, .Cols - 4) = 0
                .TextMatrix(i, .Cols - 3) = keyDocentry
                .TextMatrix(i, .Cols - 2) = iLinenum
                .TextMatrix(i, .Cols - 1) = 1
                
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False

                oRsDetail.MoveNext
                iLinenum = iLinenum + 1
            Loop
            .TextMatrix(1, 7) = 1
            .TextMatrix(1, 8) = 2
            .TextMatrix(1, 0) = -1
            .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
            lblgede = "Level (" & .TextMatrix(1, 1) & ") " & .TextMatrix(1, 2)
            
        End If

    oKon.Close
    If istatus = DataBaru And oGrid1.Rows > 1 Then
        GridModul.ClearGridDetail ogrid2
        SaveDataOgrid1_new
    End If
  
    End With
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub


Public Sub oInsertKartudetail1()
Dim irow As Integer
sQuery = "insert into master_kartu_kelas_detail1"
sQuery = sQuery & "("
sQuery = sQuery & "docentry,"
sQuery = sQuery & "linenum,"
sQuery = sQuery & "kodelevelno,"
sQuery = sQuery & "kodelevel,"
sQuery = sQuery & "kodegroup,"
sQuery = sQuery & "aktif,"
sQuery = sQuery & "status,"
sQuery = sQuery & "audituser,"
sQuery = sQuery & "auditdate"
sQuery = sQuery & ")"
sQuery = sQuery & "values"
sQuery = sQuery & "('"
sQuery = sQuery & sdocentry & "','"
sQuery = sQuery & slinenum & "','"
sQuery = sQuery & skodelevelno & "','"
sQuery = sQuery & skodelevel & "','"
sQuery = sQuery & skodegroup & "','"
sQuery = sQuery & saktif & "','"
sQuery = sQuery & sstatus & "','"
sQuery = sQuery & saudituser & "','"
sQuery = sQuery & sauditdate & "'"
sQuery = sQuery & ")"
End Sub

Public Sub oUpdateKartudetail1()
sQuery = "update  master_kartu_kelas_detail1"
sQuery = sQuery & " set "
sQuery = sQuery & "kodelevelno='" & skodelevelno & "',"
sQuery = sQuery & "kodelevel='" & skodelevel & "',"
sQuery = sQuery & "kodegroup='" & skodegroup & "',"
sQuery = sQuery & "aktif='" & 2 & "',"
sQuery = sQuery & "status='" & sstatus & "',"
sQuery = sQuery & "audituser='" & saudituser & "',"
sQuery = sQuery & "auditdate='" & sauditdate & "'"
sQuery = sQuery & " where docentry='" & sdocentry & "' and "
sQuery = sQuery & " linenum='" & slinenum & "' and aktif='1'"

End Sub



Public Sub oSaveKartuDetail1()
Dim oKon As New ADODB.Connection
Dim oRsDetail As New ADODB.Recordset
Dim irow As Integer

If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
With oGrid1
For irow = 1 To .Rows - 1
    sdocentry = .TextMatrix(irow, .Cols - 3)
    slinenum = .TextMatrix(irow, .Cols - 2)
    skodelevelno = .TextMatrix(irow, 1)
    skodelevel = .TextMatrix(irow, 3)
    skodegroup = .TextMatrix(irow, 4)
    saktif = .TextMatrix(irow, .Cols - 5)
    sstatus = .TextMatrix(irow, .Cols - 4)
    Select Case .TextMatrix(irow, .Cols - 1)
    Case 1
        oKon.Execute ("update master_kartu_kelas_detail1 set aktif='0' where docentry='" & sdocentry & "' and aktif='1'")
        oInsertKartudetail1
        Set oRsDetail = oKon.Execute(sQuery)
    Case 2
        If .TextMatrix(irow, 0) = -1 Then
        oKon.Execute ("update master_kartu_kelas_detail1 set aktif='1' where docentry='" & sdocentry & "' and linenum=" & slinenum)
        oKon.Execute ("update master_kartu_kelas_detail1 set aktif='2' where aktif='1' and docentry='" & sdocentry & "' and not linenum=" & slinenum)
        End If
'        oUpdateKartudetail1
'        Set oRsDetail = oKon.Execute(sQuery)
    End Select
    
Next
End With
    
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Kartu Detail1"
End Sub

Public Sub oInsertKartuDetail2()
sQuery = "insert into master_kartu_kelas_detail2"
sQuery = sQuery & "(docentry, "
sQuery = sQuery & "baselinenum, "
sQuery = sQuery & "linenum, "
sQuery = sQuery & "kodelevelnodetail, "
sQuery = sQuery & "tgltest, "
sQuery = sQuery & "waktupengerjaan, "
sQuery = sQuery & "jawabanbenar, "
sQuery = sQuery & "titikpangkal, "
sQuery = sQuery & "kelompok, "
sQuery = sQuery & "aktif, "
sQuery = sQuery & "STATUS, "
sQuery = sQuery & "audituser, "
sQuery = sQuery & "auditdate)"
sQuery = sQuery & " values ('"
sQuery = sQuery & sdocentry & "','"
sQuery = sQuery & sbaselinenum & "','"
sQuery = sQuery & slinenum & "','"
sQuery = sQuery & skodelevelnodetail & "','"
sQuery = sQuery & stgltest & "','"
sQuery = sQuery & swaktupengerjaan & "','"
sQuery = sQuery & sjawabanbenar & "','"
sQuery = sQuery & stitikpangkal & "','"
sQuery = sQuery & skelompok & "','"
sQuery = sQuery & saktif & "','"
sQuery = sQuery & sstatus & "','"
sQuery = sQuery & saudituser & "','"
sQuery = sQuery & sauditdate & "')"
End Sub
Public Sub oUpdateKartuDetail2()
sQuery = "update master_kartu_kelas_detail2 set "
sQuery = sQuery & "kodelevelnodetail='" & skodelevelnodetail & "',"
sQuery = sQuery & "tgltest='" & stgltest & "',"
sQuery = sQuery & "waktupengerjaan='" & swaktupengerjaan & "',"
sQuery = sQuery & "jawabanbenar='" & sjawabanbenar & "',"
sQuery = sQuery & "titikpangkal='" & stitikpangkal & "',"
sQuery = sQuery & "kelompok='" & skelompok & "',"
sQuery = sQuery & "aktif='" & saktif & "',"
sQuery = sQuery & "status='" & sstatus & "',"
sQuery = sQuery & "audituser='" & saudituser & "',"
sQuery = sQuery & "auditdate='" & sauditdate & "'"
sQuery = sQuery & " where docentry='" & sdocentry & "' and "
sQuery = sQuery & "baselinenum='" & sbaselinenum & "' and "
sQuery = sQuery & "linenum='" & slinenum & "'"
End Sub
Public Sub oSaveKartuDetail2()
On Error GoTo errhandler
Dim oKon As New ADODB.Connection
Dim oRsDetail As New ADODB.Recordset
Dim irow As Integer

If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    'oKon.Execute ("update master_kartu_kelas_detail2 set aktif='0' where docentry='" & sdocentry & "' and aktif='1'")
With ogrid2
For irow = 1 To .Rows - 1
    sdocentry = .TextMatrix(irow, .Cols - 4)
    sbaselinenum = .TextMatrix(irow, .Cols - 3)
    slinenum = .TextMatrix(irow, .Cols - 2)
    skodelevelnodetail = .TextMatrix(irow, 1)
    stgltest = Format(.TextMatrix(irow, 2), "YYYY/MM/DD")
    swaktupengerjaan = .TextMatrix(irow, 3)
    sjawabanbenar = .TextMatrix(irow, 4)
    stitikpangkal = .TextMatrix(irow, 5)
    skelompok = .TextMatrix(irow, 6)
    saktif = .TextMatrix(irow, .Cols - 6)
    sstatus = .TextMatrix(irow, .Cols - 5)
    Select Case .TextMatrix(irow, .Cols - 1)
    Case 1
        oInsertKartuDetail2
        Set oRsDetail = oKon.Execute(sQuery)
    Case 2
        oUpdateKartuDetail2
        Set oRsDetail = oKon.Execute(sQuery)
    End Select
    
Next
End With
    
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Kartu Detail1"
End Sub

Public Sub ShowGrid2_newdata(keyDocentry As Integer, keyBaseLine As Integer, keyMulai As Integer, keySelesai As Integer)
On Error GoTo errhandler
'    Dim oKon As New ADODB.Connection
'    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim iPosAktif As Integer
      
    With ogrid2

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid2
        '.ColHidden(.Cols - 1) = True
            Dim i As Double
            Dim sLinemum As Integer
            iPosAktif = 1
            For i = keyMulai To keySelesai
                    sLinemum = sLinemum + 1
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = 0
                    .TextMatrix(.Rows - 1, 1) = i
                    .TextMatrix(.Rows - 1, 2) = ""
                    .TextMatrix(.Rows - 1, 3) = ""
                    .TextMatrix(.Rows - 1, 4) = ""
                    .TextMatrix(.Rows - 1, 5) = ""
                    .TextMatrix(.Rows - 1, 6) = ""
                    .TextMatrix(.Rows - 1, 7) = 0
                    .TextMatrix(.Rows - 1, 8) = 0

                    .TextMatrix(.Rows - 1, .Cols - 4) = keyDocentry
                    .TextMatrix(.Rows - 1, .Cols - 3) = keyBaseLine
                    .TextMatrix(.Rows - 1, .Cols - 2) = sLinemum
                    .TextMatrix(.Rows - 1, .Cols - 1) = 1
                     .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = False
                    
            Next
            If .row = 0 Then Exit Sub
            .Select 1, 0
            iPosAktif = .row
            .Cell(flexcpBackColor, iPosAktif, 0, , .Cols - 1) = vbGreen
            .TextMatrix(iPosAktif, 0) = -1
            '.Select iPosAktif, 0
            If .TextMatrix(iPosAktif, 0) = -1 Then
                Text1(6).Enabled = True
                Text1(7).Enabled = True
                Text1(8).Enabled = True
                Text1(9).Enabled = True
            Else
                Text1(6).Enabled = False
                Text1(7).Enabled = False
                Text1(8).Enabled = False
                Text1(9).Enabled = False
            End If
            Text1(5) = ogrid2.TextMatrix(iPosAktif, 1)
            Text1(6) = ogrid2.TextMatrix(iPosAktif, 2)
            Text1(7) = ogrid2.TextMatrix(iPosAktif, 3)
            Text1(8) = ogrid2.TextMatrix(iPosAktif, 4)
            Text1(9) = ogrid2.TextMatrix(iPosAktif, 5)
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        
    End With

    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Function SaveDataDetail2_new() As Boolean
On Error GoTo errhandler
Dim i  As Integer
Dim oKon As New ADODB.Connection
Dim oRsDetail As New ADODB.Recordset
Dim iLevelno As Integer
Dim inolvlmulai As Integer
Dim inolvlselesai As Integer
Dim iLinenum As Integer
If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Update master_kartu_kelas_detail1 set aktif=2 where docentry=" & ToNumber(Label3.Caption)
    Set oRsDetail = oKon.Execute(sQuery)
With oGrid1
For i = 1 To .Rows - 1

            sQuery = "insert into  master_kartu_kelas_detail1 ("
            sQuery = sQuery & "docentry,"
            sQuery = sQuery & "linenum,"
            sQuery = sQuery & "kodelevelno,"
            sQuery = sQuery & "kodelevel,"
            sQuery = sQuery & "kodegroup,"
            sQuery = sQuery & "aktif,"
            sQuery = sQuery & "status,"
            sQuery = sQuery & "audituser,"
            sQuery = sQuery & "auditdate)"
            sQuery = sQuery & " values ('"
            sQuery = sQuery & .TextMatrix(i, .Cols - 3) & "','"
            sQuery = sQuery & .TextMatrix(i, .Cols - 2) & "','"
            sQuery = sQuery & .TextMatrix(i, 1) & "','"
            sQuery = sQuery & .TextMatrix(i, 3) & "','"
            sQuery = sQuery & .TextMatrix(i, 4) & "','"
            sQuery = sQuery & .TextMatrix(i, .Cols - 5) & "','"
            sQuery = sQuery & .TextMatrix(i, .Cols - 4) & "','"
            sQuery = sQuery & MenuFrm.sUserID & "','"
            sQuery = sQuery & Format(Now(), "YYYY/MM/DD") & "')"
            Set oRsDetail = oKon.Execute(sQuery)
            '----- simpan kartu detail2
            
                
                sQuery = "Select nolvlmulai, nolvlselesai from master_pelajaran_level_detail where "
                sQuery = sQuery & " kodegroup='" & .TextMatrix(i, 4) & "' and "
                sQuery = sQuery & " kodelevel='" & .TextMatrix(i, 3) & "' and "
                sQuery = sQuery & " kodelevelno='" & .TextMatrix(i, 1) & "'"
                Set oRsDetail = oKon.Execute(sQuery)
                inolvlmulai = oRsDetail("nolvlmulai")
                inolvlselesai = oRsDetail("nolvlselesai")
                For iLevelno = inolvlmulai To inolvlselesai
                    iLinenum = iLinenum + 1
                        sdocentry = .TextMatrix(i, .Cols - 3)
                        sbaselinenum = .TextMatrix(i, .Cols - 2)
                        slinenum = iLinenum
                        skodelevelnodetail = iLevelno
                        stgltest = ""
                        swaktupengerjaan = ""
                        sjawabanbenar = ""
                        stitikpangkal = ""
                        skelompok = ""
                        saktif = "0"
                        sstatus = "0"
                        oInsertKartuDetail2
                        Set oRsDetail = oKon.Execute(sQuery)
                Next
Next
End With

oKon.Close
SaveDataDetail2_new = True
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "Save DataDetail2_new"
End Function
Function SaveDataOgrid1_new() As Boolean
Dim ires As Integer
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        SaveDataOgrid1_new = False
       If SaveDataDetail2_new Then
             istatus = Normal
             ShowGrid1 sdocentry, skodelevel
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringLevelKelas
             MenuFrm.Toolbar1.Buttons(btm_execut).Enabled = True
             SaveDataOgrid1_new = True
        End If
    End If
    
End Function
Public Sub Execution()
On Error GoTo errhandler

Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\DaftarJelangTPP.Rpt"

Dim sKriteria As String

'sKriteria = " where nokursus  between '" & ogrid3.TextMatrix(ogrid3.row, 3) & "' and '" & ogrid3.TextMatrix(ogrid3.row, 3) & "'"

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vmaster_jelang_tpp_rpt vmaster_jelang_tpp_rpt1"
'
'
Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "cmpnyname" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(1) = "address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
Me.cr1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Public Sub ShowGridKelas(snoidsiswaQ As String, sKriteria As String, sOrderBy As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
      
    sQuery = "SELECT docentry,nokursus,noidsiswa,nmlengkap, tglmulai,materi,titikpangkal2 AS titikpangkal ,"
    sQuery = sQuery & " namapembimbing,pelajaran,kodegroup FROM  vmaster_kelas_rpt2 "
    sQuery = sQuery & " WHERE pelajaran='" & snoidsiswaQ & "' "
    sQuery = sQuery & " and " & sKriteria & " order by " & sOrderBy & ""
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid3
        
        GridModul.ClearGridDetail ogrid3
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double

            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                
                .TextMatrix(i, 0) = oRsDetail("docentry")
                .TextMatrix(i, 1) = oRsDetail("noidsiswa")
                .TextMatrix(i, 2) = oRsDetail("nmlengkap")
                .TextMatrix(i, 3) = RTrim(oRsDetail("nokursus"))
                .TextMatrix(i, 4) = ToText(oRsDetail("tglmulai"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("materi"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("namapembimbing"))
                                
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                oRsDetail.MoveNext
            Loop
            End If
            If i = 0 Then
                Label3.Caption = 0
            Else
                .Select 1, 0
                Label3.Caption = .TextMatrix(1, 0)
            End If
    End With
    
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub oGetMateriku(sDocentryQ As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    If oKon.State = 1 Then oKon.Close
        oKon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "SELECT a.docentry,a.kodegroup,b.keterangan AS materi, a.kodelevel ,c.namalevel"
        sQuery = sQuery & " FROM master_kartu_kelas_detail1 a "
        sQuery = sQuery & " INNER JOIN master_default_pelajaran b ON a.kodegroup=b.kodegroup AND a.aktif='1' AND docentry=" & sDocentryQ
        sQuery = sQuery & " INNER JOIN master_pelajaran_level c ON a.kodegroup=c.kodegroup AND a.kodelevel=c.kodelevel"
        Set oRsDetail = oKon.Execute(sQuery)
    If Not oRsDetail.EOF Then
        skodegroup = oRsDetail("kodegroup")
        Text1(3) = oRsDetail("materi")
        Text1(4) = oRsDetail("kodelevel")
        Label4.Caption = oRsDetail("namalevel")
        
    End If
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "oGetMateriku"
End Sub

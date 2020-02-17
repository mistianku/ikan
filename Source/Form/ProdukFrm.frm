VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form ProdukFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Produk Form"
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   7680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Detail Produk"
            Key             =   "keyDetail"
            Object.ToolTipText     =   "Infromasi Terkait dengan Detail Produk"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Diskon"
            Key             =   "keyDiskon"
            Object.ToolTipText     =   "Informasi Diskon Per Produk"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fee"
            Key             =   "keyFee"
            Object.ToolTipText     =   "Informasi Fee Produk"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Harga Satuan"
            Key             =   "keyHargaSatuan"
            Object.ToolTipText     =   "Informasi Harga Satuan Per Produk"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inventori"
            Key             =   "keyInventori"
            Object.ToolTipText     =   "Informasi Tentang Stok Barang di Masing2 Gudang"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   10200
         TabIndex        =   11
         Top             =   120
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
         Index           =   1
         Left            =   2340
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   9195
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
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2775
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukFrm.frx":0000
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
         Caption         =   "Nama Produk"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Produk"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Harga Produk"
      Height          =   6255
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   5655
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   11295
         _cx             =   19923
         _cy             =   9975
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ProdukFrm.frx":001C
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
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Fee Produk"
      Height          =   6255
      Index           =   6
      Left            =   120
      TabIndex        =   47
      Top             =   1320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   5655
         Index           =   3
         Left            =   240
         TabIndex        =   48
         Top             =   360
         Width           =   11295
         _cx             =   19923
         _cy             =   9975
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ProdukFrm.frx":00A6
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
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Diskon"
      Height          =   6255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   5655
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   11295
         _cx             =   19923
         _cy             =   9975
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ProdukFrm.frx":0127
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
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Produk Detal"
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   11775
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   2340
         TabIndex        =   46
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
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
         Index           =   14
         Left            =   5640
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   1440
         Width           =   5895
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
         Left            =   5640
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   1080
         Width           =   5895
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
         Left            =   5640
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   720
         Width           =   5895
      End
      Begin VB.Frame frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Satuan Barang"
         Height          =   1695
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   11535
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
            Left            =   5520
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   1080
            Width           =   5885
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
            Left            =   5520
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   720
            Width           =   5885
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
            Left            =   5520
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   360
            Width           =   5885
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
            Left            =   3240
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   1080
            Width           =   1790
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
            Left            =   3240
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   720
            Width           =   1790
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
            Left            =   3240
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   360
            Width           =   1790
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
            Index           =   8
            Left            =   2240
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1080
            Width           =   975
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
            Index           =   7
            Left            =   2240
            TabIndex        =   28
            Text            =   "Text1"
            Top             =   720
            Width           =   975
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
            Index           =   6
            Left            =   2240
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   360
            Width           =   975
         End
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   255
            Index           =   5
            Left            =   5040
            TabIndex        =   33
            Top             =   720
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            MouseIcon       =   "ProdukFrm.frx":01AB
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
            Index           =   6
            Left            =   5040
            TabIndex        =   34
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            MouseIcon       =   "ProdukFrm.frx":01C7
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
            Left            =   5040
            TabIndex        =   35
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            MouseIcon       =   "ProdukFrm.frx":01E3
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
            Caption         =   "UoM3"
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "UoM2"
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "UoM1"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   360
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
         Index           =   5
         Left            =   2340
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1440
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
         Index           =   4
         Left            =   2340
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1080
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
         Index           =   3
         Left            =   2340
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   720
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
         Index           =   2
         Left            =   2340
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
         Width           =   2775
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   20
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukFrm.frx":01FF
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
         Left            =   5160
         TabIndex        =   21
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukFrm.frx":021B
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
         Left            =   5160
         TabIndex        =   22
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukFrm.frx":0237
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
         Caption         =   "Tanggal Register"
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   36
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Fungsi"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Kategori"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Brand"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Barcode"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Inventori"
      Height          =   6255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   5655
         Index           =   2
         Left            =   240
         TabIndex        =   45
         Top             =   360
         Width           =   11295
         _cx             =   19923
         _cy             =   9975
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ProdukFrm.frx":0253
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
End
Attribute VB_Name = "ProdukFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sInsertProdukDiskon As String
Dim sInsertProdukHarga As String
Dim sInsertProdukInventori As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm
Dim skodeproduk As String
Dim snamaproduk As String
Dim saktif As String
Dim skodebrand As String
Dim skodekategori As String
Dim skodefungsi As String
Dim skodebarcode As String
Dim suom1 As Integer
Dim suom2 As Integer
Dim sumo3 As Integer
Dim suom1sat As String
Dim suom2sat As String
Dim suom3sat As String
Dim sregisterdate As String
Dim saudituser As String
Dim sauditdate  As String

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from vmaster_produk where kodeproduk='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm
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
    sQuery = "Select *  from vmaster_produk order by kodeproduk asc limit 1"
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
    sQuery = "Select  *  from vmaster_produk where kodeproduk >'" & Text1(0).text & "' order by kodeproduk asc limit 1"
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
    sQuery = "Select  *  from vmaster_produk where kodeproduk<'" & Text1(0).text & "' order by kodeproduk desc limit 1"
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
    sQuery = "Select *  from vmaster_produk order by kodeproduk desc limit 1 "
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
             FindData Text1(0)
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm
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
        If istatus = StatusForm.DataBaru Then
'            oCon.Execute sInsertProdukDiskon
'            oCon.Execute sInsertProdukHarga
'            oCon.Execute sInsertProdukInventori
'            oCon.Execute "Call sp_master_produk_harga_customer_insert('" & skodeproduk & "','" & MenuFrm.sUserID & "')"
        Else
            SimpanGrid1 Text1(0)
            SimpanGrid2 Text1(0)
            SimpanGrid3 Text1(0)
            
            
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    FlatDatePicker1.value = Now()
    Text1(6) = 1
    Text1(7) = 1
    Text1(8) = 1
    GridModul.ClearGridDetail oGrid1(0)
    GridModul.ClearGridDetail oGrid1(1)
    GridModul.ClearGridDetail oGrid1(2)
    GridModul.ClearGridDetail oGrid1(3)
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skodeproduk = ToText(Text1(0).text)
    snamaproduk = ToText(Text1(1).text)
    
    If Check1(0).value = 0 Then
        saktif = "0"
    Else
        saktif = "1"
    End If
        
        skodebrand = ToText(Text1(3).text)
        skodekategori = ToText(Text1(4))
        skodefungsi = ToText(Text1(5))
        skodebarcode = ToText(Text1(2))
        suom1 = ToNumber(Text1(6))
        suom2 = ToNumber(Text1(7))
        sumo3 = ToNumber(Text1(8))
        suom1sat = ToText(Text1(9))
        suom2sat = ToText(Text1(10))
        suom3sat = ToText(Text1(11))
        sregisterdate = Format(FlatDatePicker1.value, "YYYY-MM-DD")

sUpdate = "call sp_master_produk_update('"
sUpdate = sUpdate & skodeproduk & "','"
sUpdate = sUpdate & snamaproduk & "','"
sUpdate = sUpdate & saktif & "','"
sUpdate = sUpdate & skodebrand & "','"
sUpdate = sUpdate & skodekategori & "','"
sUpdate = sUpdate & skodefungsi & "','"
sUpdate = sUpdate & skodebarcode & "','"
sUpdate = sUpdate & suom1 & "','"
sUpdate = sUpdate & suom2 & "','"
sUpdate = sUpdate & sumo3 & "','"
sUpdate = sUpdate & suom1sat & "','"
sUpdate = sUpdate & suom2sat & "','"
sUpdate = sUpdate & suom3sat & "','"
sUpdate = sUpdate & sregisterdate & "','"
sUpdate = sUpdate & MenuFrm.sUserID & "')"

sInsert = Replace(sUpdate, "update", "insert")
sDelete = Replace(sUpdate, "update", "delete")

    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsMasterProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(3) = oBrowse.YangDipilih
        Text1(12) = oBrowse.Keterangan
        Text1(3).SetFocus
    End If
Case 2
    oBrowse.ShowFinder BrowsCategory, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(4) = oBrowse.YangDipilih
        Text1(13) = oBrowse.Keterangan
        Text1(4).SetFocus
    End If
Case 3
    oBrowse.ShowFinder BrowsFunction, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(5) = oBrowse.YangDipilih
        Text1(14) = oBrowse.Keterangan
        Text1(5).SetFocus
    End If
Case 4
    oBrowse.ShowFinder BrowsSatuanProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(9) = oBrowse.YangDipilih
        Text1(15) = oBrowse.Keterangan
        Text1(9).SetFocus
    End If
Case 5
    oBrowse.ShowFinder BrowsSatuanProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(10) = oBrowse.YangDipilih
        Text1(16) = oBrowse.Keterangan
        Text1(10).SetFocus
    End If
Case 6
    oBrowse.ShowFinder BrowsSatuanProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(11) = oBrowse.YangDipilih
        Text1(17) = oBrowse.Keterangan
        Text1(11).SetFocus
    End If
End Select


Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Produk"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukFrm

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(3).Top
BrowseUserID(1).Height = Text1(3).Height
BrowseUserID(1).Left = Text1(3).Left + Text1(3).Width

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width

BrowseUserID(3).Top = Text1(5).Top
BrowseUserID(3).Height = Text1(5).Height
BrowseUserID(3).Left = Text1(5).Left + Text1(5).Width

BrowseUserID(4).Top = Text1(9).Top
BrowseUserID(4).Height = Text1(9).Height
BrowseUserID(4).Left = Text1(9).Left + Text1(9).Width

BrowseUserID(5).Top = Text1(10).Top
BrowseUserID(5).Height = Text1(10).Height
BrowseUserID(5).Left = Text1(10).Left + Text1(10).Width

BrowseUserID(6).Top = Text1(11).Top
BrowseUserID(6).Height = Text1(11).Height
BrowseUserID(6).Left = Text1(11).Left + Text1(11).Width

End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me

oFormatCheckList 1, Me
Frame1(4).ZOrder
cleardata
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("kodeproduk")
    KodeUserAksesTemp = oRs("kodeproduk")
    Text1(0).Locked = True
    Text1(1).text = oRs("namaproduk")
        
        Check1(0).value = IIf(oRs("Aktif") = "1", 1, 0)
        Text1(0) = oRs("kodeproduk")
        Text1(1) = oRs("namaproduk")
        
        Text1(3) = oRs("kodebrand")
        Text1(4) = oRs("kodekategori")
        Text1(5) = oRs("kodefungsi")
        Text1(2) = oRs("kodebarcode")
        Text1(6) = oRs("uom1")
        Text1(7) = oRs("uom2")
        Text1(8) = oRs("umo3")
        Text1(9) = oRs("uom1sat")
        Text1(10) = oRs("uom2sat")
        Text1(11) = oRs("uom3sat")
        FlatDatePicker1.value = oRs("registerdate")
        
        Text1(12) = ToText(oRs("namabrand"))
        Text1(13) = ToText(oRs("namakategori"))
        Text1(14) = ToText(oRs("namafungsi"))
        Text1(15) = ToText(oRs("namasatuan1"))
        Text1(16) = ToText(oRs("namasatuan2"))
        Text1(17) = ToText(oRs("namasatuan3"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    
    ShowGrid1 ToText(Text1(0).text)
    ShowGrid2 Text1(0)
    ShowGrid3 Text1(0)
    ShowGrid4 Text1(0)
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

Private Sub ogrid1_CellChanged(Index As Integer, ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = oGrid1(Index).Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, oGrid1(Index), istatus
End Sub

Private Sub oGrid1_Click(Index As Integer)
Select Case Index
Case 0
    With oGrid1(Index)
        If .col = 2 Then
            .EditCell
        End If
    End With
Case 1
    With oGrid1(Index)
        If .col = 2 Then
            .EditCell
        End If
    End With
Case 3
    With oGrid1(Index)
        If .col = 2 Then
            .EditCell
        End If
    End With
End Select
End Sub

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "keyDetail"
            Frame1(0).ZOrder   'Picture1(0).ZOrder
Case "keyDiskon"
            Frame1(3).ZOrder   'Picture1(0).ZOrder
Case "keyFee"
            Frame1(6).ZOrder   'Picture1(0).ZOrder
Case "keyHargaSatuan"
            Frame1(4).ZOrder   'Picture1(0).ZOrder
Case "keyInventori"
            Frame1(2).ZOrder   'Picture1(0).ZOrder
End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "Informasi Produk"
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

Public Sub ShowGrid1(keyKodeProduk As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "select * from vmaster_produk_diskon "
    sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' order by kodediskon asc"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(0)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1(0)
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodediskon"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namadiskon"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("diskon"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub ShowGrid2(keyKodeProduk As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "select * from vmaster_produk_harga "
    sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' order by kodeharga asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(1)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1(1)
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodeharga"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namadiskon"))
                .TextMatrix(i, 2) = ToNumber(oRsDetail("harga"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub

Public Sub ShowGrid3(keyKodeProduk As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "select * from vmaster_inventori "
    sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' order by kodegudang asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(2)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1(2)
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodegudang"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namagudang"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("stock"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub ShowGrid4(keyKodeProduk As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "select * from vmaster_produk_fee "
    sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' order by kodediskon asc"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(3)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1(3)
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodediskon"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namadiskon"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("diskon"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub SimpanGrid1(keyKodeProduk As String)
On Error GoTo errhandler


    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With oGrid1(0)
        For irow = 1 To .Rows - 1
            sQuery = "Update master_produk_diskon set diskon='" & toNumberIndonesia(.TextMatrix(irow, 2)) & "'"
            sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' and kodediskon='" & .TextMatrix(irow, 0) & "'"
            If .TextMatrix(irow, .Cols - 1) = "2" Then
                oKon.Execute (sQuery)
                If is_update Then
                    If MsgBox("Update Diskon Supplier", vbYesNo, "Pesan Update Fee Customer") = vbYes Then
                        is_update = False
                    End If
                End If
                If Not is_update Then
                    oCon.Execute "Call sp_master_produk_harga_supplier_update_by_diskon('" & keyKodeProduk & "','" & MenuFrm.sUserID & "')"
                End If
            End If
        Next
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub

Public Sub SimpanGrid2(keyKodeProduk As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With oGrid1(1)
        For irow = 1 To .Rows - 1
            sQuery = "Update master_produk_harga set harga='" & toNumberIndonesia(.TextMatrix(irow, 2)) & "'"
            sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' and kodeharga='" & .TextMatrix(irow, 0) & "'"
            If .TextMatrix(irow, .Cols - 1) = "2" Then
                oKon.Execute (sQuery)
                If Not MenuFrm.sKodeHargaCost = .TextMatrix(irow, 0) Then
                
                    If is_update Then
                        If MsgBox("Update Harga Produk Per Customer", vbYesNo, "Pesan Update Harga Customer") = vbYes Then
                            is_update = False
                        End If
                    End If
                    If Not is_update Then
                        oCon.Execute "Call sp_master_produk_harga_customer_update_by_harga('" & keyKodeProduk & "','" & MenuFrm.sUserID & "')"
                    End If
                Else
                    If MsgBox("Update HPP untuk 1 Bulan Transaksi Kebelakang ?", vbQuestion + vbYesNo, "Update Transaksi HPP") = vbYes Then
                       oCon.Execute "Call sp_update_hpp_by_master_produk('" & keyKodeProduk & "','" & toNumberIndonesia(.TextMatrix(irow, 2)) & "','" & MenuFrm.sUserID & "')"
                    End If
                End If
                           
            End If
        Next
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub
Public Sub SimpanGrid3(keyKodeProduk As String)
On Error GoTo errhandler


    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With oGrid1(3)
        For irow = 1 To .Rows - 1
            sQuery = "Update master_produk_fee set diskon='" & toNumberIndonesia(.TextMatrix(irow, 2)) & "'"
            sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' and kodediskon='" & .TextMatrix(irow, 0) & "'"
            If .TextMatrix(irow, .Cols - 1) = "2" Then
                oKon.Execute (sQuery)
                If is_update Then
                    If MsgBox("Update Fee Customer", vbYesNo, "Pesan Update Fee Customer") = vbYes Then
                        is_update = False
                    End If
                End If
                If Not is_update Then
                    oCon.Execute "Call sp_master_produk_harga_customer_update_by_fee('" & keyKodeProduk & "','" & MenuFrm.sUserID & "')"
                End If
            End If
        Next
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Fee"
End Sub

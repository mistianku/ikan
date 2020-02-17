VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form CustomerFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Customer Form"
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
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
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
         Index           =   27
         Left            =   5400
         TabIndex        =   87
         Text            =   "Text1"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Lock Sales"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   9840
         TabIndex        =   79
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   9840
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
         Left            =   2460
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
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
         Left            =   2460
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MouseIcon       =   "CustomerFrm.frx":0000
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
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   9840
         TabIndex        =   80
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         Locked          =   -1  'True
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Customer"
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Customer"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   7800
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Placement       =   1
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info Master Customer"
            Key             =   "keyCustomer"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Master Harga"
            Key             =   "keyProduct"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Info PT"
            Key             =   "keyCompany"
            Object.ToolTipText     =   "Info PTyang digunakan untuk Faktur"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Insentif Sales"
            Key             =   "keyInsentif"
            Object.ToolTipText     =   "Info Setting Insentif "
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Master Harga"
      Height          =   6495
      Index           =   7
      Left            =   120
      TabIndex        =   51
      Top             =   1320
      Width           =   12015
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Cari"
         Height          =   315
         Left            =   4440
         TabIndex        =   59
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Otomatis"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10440
         TabIndex        =   58
         Top             =   240
         Width           =   1455
      End
      Begin VSDFLATS.FlatComboBox searchby 
         Height          =   285
         Left            =   2460
         TabIndex        =   55
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
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
         MouseIcon       =   "CustomerFrm.frx":001C
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
         Left            =   2460
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   720
         Width           =   7155
      End
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   4695
         Index           =   0
         Left            =   360
         TabIndex        =   57
         Top             =   1200
         Width           =   11535
         _cx             =   20346
         _cy             =   8281
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   12632256
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"CustomerFrm.frx":0038
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kunci Pencarian"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   23
         Left            =   9480
         TabIndex        =   72
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kunci Pencarian"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   22
         Left            =   1850
         TabIndex        =   71
         Top             =   6000
         Width           =   7575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kunci Pencarian"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   21
         Left            =   360
         TabIndex        =   70
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Kunci Pencarian"
         Height          =   315
         Index           =   18
         Left            =   360
         TabIndex        =   54
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cari Berdasarkan"
         Height          =   315
         Index           =   17
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Master Customer"
      Height          =   6495
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   12015
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   3480
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
            Index           =   2
            Left            =   2340
            TabIndex        =   46
            Text            =   "Text1"
            Top             =   240
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
            Index           =   3
            Left            =   2340
            TabIndex        =   45
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
            Index           =   4
            Left            =   2340
            TabIndex        =   44
            Text            =   "Text1"
            Top             =   960
            Width           =   2955
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
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   1320
            Width           =   2955
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
            Left            =   5400
            TabIndex        =   42
            Text            =   "Text1"
            Top             =   1320
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Alamat"
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Kota"
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   49
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Telp/Faximale"
            Height          =   315
            Index           =   4
            Left            =   240
            TabIndex        =   48
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Faximale"
            Height          =   315
            Index           =   5
            Left            =   8400
            TabIndex        =   47
            Top             =   1320
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   1815
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   240
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
            Index           =   26
            Left            =   7980
            TabIndex        =   82
            Text            =   "Text1"
            Top             =   1320
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
            Index           =   25
            Left            =   9840
            TabIndex        =   81
            Text            =   "Text1"
            Top             =   1320
            Width           =   1875
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
            Index           =   20
            Left            =   4200
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   1320
            Width           =   1875
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
            TabIndex        =   60
            Text            =   "Text1"
            Top             =   1320
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
            Index           =   7
            Left            =   2340
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   240
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
            Index           =   8
            Left            =   2340
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   960
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
            Index           =   9
            Left            =   2340
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   600
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
            Index           =   10
            Left            =   4200
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   240
            Width           =   5235
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
            Left            =   4200
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   960
            Width           =   5235
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
            Left            =   4200
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   600
            Width           =   5235
         End
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   285
            Index           =   1
            Left            =   3720
            TabIndex        =   35
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":0166
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
            Height          =   285
            Index           =   2
            Left            =   3720
            TabIndex        =   36
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":0182
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
            Height          =   285
            Index           =   3
            Left            =   3720
            TabIndex        =   37
            Top             =   600
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":019E
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
            Height          =   285
            Index           =   5
            Left            =   3720
            TabIndex        =   63
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":01BA
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
            Height          =   285
            Index           =   8
            Left            =   9360
            TabIndex        =   83
            Top             =   1320
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":01D6
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
            Caption         =   "Kode Purchasing"
            Height          =   315
            Index           =   25
            Left            =   6120
            TabIndex        =   84
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Kode Salesman"
            Height          =   315
            Index           =   19
            Left            =   240
            TabIndex        =   61
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Kode Gudang"
            Height          =   315
            Index           =   8
            Left            =   240
            TabIndex        =   40
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Kode Harga"
            Height          =   315
            Index           =   9
            Left            =   240
            TabIndex        =   39
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Kode Diskon"
            Height          =   315
            Index           =   10
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "PIC"
         Height          =   1095
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   5280
         Width           =   11775
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Tampil Info Company"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   8040
            TabIndex        =   78
            Top             =   600
            Width           =   3375
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Model Faktur Customer"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   77
            Top             =   600
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
            Index           =   24
            Left            =   8520
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   240
            Width           =   2895
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
            Index           =   23
            Left            =   6900
            TabIndex        =   73
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
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
            TabIndex        =   25
            Text            =   "Text1"
            Top             =   240
            Width           =   2955
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
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   600
            Width           =   2955
         End
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   285
            Index           =   7
            Left            =   8040
            TabIndex        =   74
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":01F2
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
            Caption         =   "Kode Area"
            Height          =   315
            Index           =   24
            Left            =   5400
            TabIndex        =   75
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Nama"
            Height          =   315
            Index           =   11
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "No.HP"
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   5400
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
            Index           =   15
            Left            =   2340
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   240
            Width           =   1335
         End
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   285
            Index           =   4
            Left            =   3720
            TabIndex        =   21
            Top             =   240
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":020E
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
            Caption         =   "No.ID.Siswa"
            Height          =   315
            Index           =   6
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   1455
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2040
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
            Index           =   22
            Left            =   4200
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   960
            Width           =   5235
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
            Index           =   21
            Left            =   2340
            TabIndex        =   64
            Text            =   "Text1"
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Kredit"
            Height          =   315
            Index           =   2
            Left            =   5040
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Transfer"
            Height          =   315
            Index           =   1
            Left            =   3720
            TabIndex        =   12
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tunai"
            Height          =   315
            Index           =   0
            Left            =   2400
            TabIndex        =   13
            Top             =   240
            Width           =   1815
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
            Index           =   16
            Left            =   2340
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
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
            Index           =   17
            Left            =   6540
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   600
            Width           =   1335
         End
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   285
            Index           =   6
            Left            =   3720
            TabIndex        =   65
            Top             =   960
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            MouseIcon       =   "CustomerFrm.frx":022A
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
            Caption         =   "Kode Fee"
            Height          =   315
            Index           =   20
            Left            =   240
            TabIndex        =   66
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Jenis Bayar"
            Height          =   315
            Index           =   12
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Tempo Bayar"
            Height          =   315
            Index           =   13
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "PPN"
            Height          =   315
            Index           =   14
            Left            =   4440
            TabIndex        =   16
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Hari"
            Height          =   315
            Index           =   15
            Left            =   3720
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            Height          =   315
            Index           =   16
            Left            =   7920
            TabIndex        =   14
            Top             =   600
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Insentif Sales"
      Height          =   6495
      Index           =   9
      Left            =   120
      TabIndex        =   85
      Top             =   1320
      Width           =   12015
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   5175
         Index           =   2
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   9128
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   12632256
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"CustomerFrm.frx":0246
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
      Caption         =   "Info PT"
      Height          =   6495
      Index           =   8
      Left            =   120
      TabIndex        =   68
      Top             =   1320
      Width           =   12015
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   5175
         Index           =   1
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   10815
         _cx             =   19076
         _cy             =   9128
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   12632256
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"CustomerFrm.frx":02F5
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
Attribute VB_Name = "CustomerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm

Dim lasposrow As Integer

Dim skodecustomer As String
Dim snamacustomer As String
Dim skodeharga As String
Dim skodediskon As String
Dim sfee As String
Dim skodegudang As String
Dim sppn As Integer
Dim sjtempo As Integer
Dim sjbayar As String
Dim salamat1 As String
Dim salamat2 As String
Dim skota As String
Dim stelp As String
Dim sFaximale As String
Dim saktif As String
Dim spic As String
Dim spichp As String
Dim saudituser As String
Dim sauditdate As String
Dim skodesalesman, skodepurchasing As String
Dim sis_invoice_spc As Integer
Dim skodearea As String
Dim sis_print_company As String
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Call sp_master_customer_mov_new('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
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
    sQuery = "Call sp_master_customer_mov_new('" & Text1(0).text & "',1)"
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
    sQuery = "Call sp_master_customer_mov_new('" & Text1(0).text & "',3)"
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
    sQuery = "Call sp_master_customer_mov_new('" & Text1(0).text & "',2)"
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
    sQuery = "Call sp_master_customer_mov_new('" & Text1(0).text & "',4)"
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
    
    
    If Text1(9) = "" Then
        MsgBox "Satuan Harga Kosong !!", vbOKOnly, "Isi Satuan Harga"
        Exit Sub
    End If
    If Text1(21) = "" Then
        MsgBox "Satuan Fee Kosong !!", vbOKOnly, "Isi Satuan Fee"
        Exit Sub
    End If
    
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             oShowGrid1
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
             
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        SimpanGrid1 Text1(0), Text1(9)
        SimpanGrid2 Text1(0)
        SimpanGrid3
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Check1(0).value = 1
    Option1(0).value = True
    Text1(16) = 0
    Text1(17) = 0
    GridModul.ClearGridDetail ogrid(0)
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skodecustomer = ToText(Text1(0).text)
    snamacustomer = ToText(Text1(1).text)
    skodesalesman = ToText(Text1(19).text)
    skodepurchasing = ToText(Text1(26).text)
    sjtempo = ToNumber(Text1(16))
    sppn = ToNumber(Text1(17))
    sjbayar = IIf(Option1(0).value = True, "1", IIf(Option1(1).value = True, "2", "3"))
    salamat1 = ToText(Text1(2))
    salamat2 = ToText(Text1(3))
    skota = (ToText(Text1(4)))
    stelp = ToText(Text1(5))
    sFaximale = ToText(Text1(6))
   
    If Check1(0).value = "1" Then
        saktif = "1"
    Else
        saktif = "0"
    End If
    
    If Check1(1).value = "1" Then
        sis_invoice_spc = "1"
    Else
        sis_invoice_spc = "0"
    End If
    
    If Check1(2).value = "1" Then
        sis_print_company = "1"
    Else
        sis_print_company = "0"
    End If
    
    
    
    skodegudang = ToText(Text1(7))
    skodediskon = ToText(Text1(8))
    sfee = ToText(Text1(21))
    skodeharga = ToText(Text1(9))
    spic = ToText(Text1(13))
    spichp = ToText(Text1(14))
    skodearea = ToText(Text1(23))
    skodearea = IIf(skodearea = "", MenuFrm.skodeareaDefault, skodearea)
'    skodesalesman= ToText(Text1(15))
    sUpdate = "call sp_master_customer_update_new('"
    sUpdate = sUpdate & skodecustomer & "','"
    sUpdate = sUpdate & snamacustomer & "','"
    sUpdate = sUpdate & skodesalesman & "','"
    sUpdate = sUpdate & skodeharga & "','"
    sUpdate = sUpdate & skodediskon & "','"
    sUpdate = sUpdate & skodegudang & "','"
    sUpdate = sUpdate & sfee & "','"
    sUpdate = sUpdate & sppn & "','"
    sUpdate = sUpdate & sjtempo & "','"
    sUpdate = sUpdate & sjbayar & "','"
    sUpdate = sUpdate & salamat1 & "','"
    sUpdate = sUpdate & salamat2 & "','"
    sUpdate = sUpdate & skota & "','"
    sUpdate = sUpdate & stelp & "','"
    sUpdate = sUpdate & sFaximale & "','"
    sUpdate = sUpdate & saktif & "','"
    sUpdate = sUpdate & spic & "','"
    sUpdate = sUpdate & spichp & "','"
    sUpdate = sUpdate & skodearea & "','"
    sUpdate = sUpdate & sis_invoice_spc & "','"
    sUpdate = sUpdate & sis_print_company & "','"
    sUpdate = sUpdate & skodepurchasing & "','"
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
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(10) = oBrowse.Keterangan
    End If
Case 2
    oBrowse.ShowFinder BrowsDiskon, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(8) = oBrowse.YangDipilih
        Text1(11) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsHarga, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(9) = oBrowse.YangDipilih
        Text1(12) = oBrowse.Keterangan
    End If
Case 5
    oBrowse.ShowFinder BrowsSalesman, "is_salesman='1'", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(19) = oBrowse.YangDipilih
        Text1(20) = oBrowse.Keterangan
    End If
Case 6
    oBrowse.ShowFinder BrowsFee, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(21) = oBrowse.YangDipilih
        Text1(22) = oBrowse.Keterangan
    End If
Case 7
    oBrowse.ShowFinder BrowsArea, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(23) = oBrowse.YangDipilih
        Text1(24) = oBrowse.Keterangan
    End If
Case 8
    oBrowse.ShowFinder BrowsSalesman, "is_purchase='1'", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(26) = oBrowse.YangDipilih
        Text1(25) = oBrowse.Keterangan
    End If
End Select

Set oBrowse = Nothing
End Sub

Private Sub Check2_Click()
If Check2.value = Checked Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
oShowGrid1
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Customer"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnCustomerFrm
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(7).Top
BrowseUserID(1).Height = Text1(7).Height
BrowseUserID(1).Left = Text1(7).Left + Text1(7).Width

BrowseUserID(2).Top = Text1(8).Top
BrowseUserID(2).Height = Text1(8).Height
BrowseUserID(2).Left = Text1(8).Left + Text1(8).Width

BrowseUserID(3).Top = Text1(9).Top
BrowseUserID(3).Height = Text1(9).Height
BrowseUserID(3).Left = Text1(9).Left + Text1(9).Width

BrowseUserID(4).Top = Text1(15).Top
BrowseUserID(4).Height = Text1(15).Height
BrowseUserID(4).Left = Text1(15).Left + Text1(15).Width

BrowseUserID(5).Top = Text1(19).Top
BrowseUserID(5).Height = Text1(19).Height
BrowseUserID(5).Left = Text1(19).Left + Text1(19).Width

BrowseUserID(6).Top = Text1(21).Top
BrowseUserID(6).Height = Text1(21).Height
BrowseUserID(6).Left = Text1(21).Left + Text1(19).Width

oFormatOption 1, Me


End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatCheckList 1, Me
cleardata
istatus = Normal
lasposrow = 1
MoveLast

Frame1(6).ZOrder
searchby.AddItem "Nama Barang"
searchby.AddItem "Kode Barang"
searchby.ListIndex = 0
If Check2.value = Checked Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = ToText(oRs("kodecustomer"))
    KodeUserAksesTemp = ToText(oRs("kodecustomer"))
    Text1(0).Locked = True
    Text1(1).text = ToText(oRs("namacustomer"))
    Text1(2) = ToText(oRs("alamat1"))
    Text1(3) = ToText(oRs("alamat2"))
    Text1(4) = ToText(oRs("kota"))
    Text1(5) = ToText(oRs("telp"))
    Text1(6) = ToText(oRs("faximale"))
    Text1(7) = ToText(oRs("kodegudang"))
    Text1(8) = ToText(oRs("kodediskon"))
    Text1(9) = ToText(oRs("kodeharga"))
    Text1(10) = ToText(oRs("namagudang"))
    Text1(11) = ToText(oRs("namadiskon"))
    Text1(12) = ToText(oRs("namaharga"))
    Text1(13) = ToText(oRs("pic"))
    Text1(14) = ToText(oRs("pichp"))
    Text1(15) = ToText(ToText(oRs("kodesalesman")))
    Text1(16) = ToText(ToText(oRs("jtempo")))
    Text1(17) = ToText(ToText(oRs("ppn")))
    Text1(19) = ToText(ToText(oRs("kodesalesman")))
    Text1(20) = ToText(ToText(oRs("namasalesman")))
    Text1(26) = ToText(ToText(oRs("kodepurchasing")))
    Text1(25) = ToText(ToText(oRs("namapurchasing")))
    Text1(21) = ToText(ToText(oRs("fee")))
    Text1(22) = ToText(ToText(oRs("namafee")))
    Text1(23) = ToText(ToText(oRs("kodearea")))
    Text1(24) = oFindByQuery("select namaarea from master_area where kodearea='" & Text1(23) & "'", DBaseConection.Modul)
    FlatDatePicker1.value = IIf(oRs("lock_start_date") = "", Now(), oRs("lock_start_date"))
    If oRs("lock_sales") = "1" Then
        Check1(3).value = 1
    Else
        Check1(3).value = 0
    End If
    
    If oRs("jbayar") = "1" Then
         Option1(0).value = True
    Else
        If oRs("jbayar") = "1" Then
            Option1(1).value = True
        Else
             Option1(2).value = True
        End If
    End If

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    
    If oRs("aktif") = "1" Then
        Check1(0).value = 1
    Else
        Check1(0).value = 0
    End If
    
    If oRs("is_invoice_spc") = "1" Then
        Check1(1).value = 1
    Else
        Check1(1).value = 0
    End If
    '
    If oRs("is_print_company") = "1" Then
        Check1(2).value = 1
    Else
        Check1(2).value = 0
    End If
    
    
    oShowGrid1
    oShowGrid2
    oShowGrid3
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

Private Sub ogrid_CellChanged(Index As Integer, ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = ogrid(Index).Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, ogrid(Index), istatus
End Sub

Private Sub ogrid_Click(Index As Integer)
Dim srow_poss As Integer
Dim irow As Integer
srow_poss = ogrid(Index).row
If Index = 0 Then
    With ogrid(0)
    Label1(21).Caption = .TextMatrix(.row, 0)
    Label1(22).Caption = .TextMatrix(.row, 5)
    Label1(23).Caption = .TextMatrix(.row, 4)
    lasposrow = .row
    Select Case .col
    Case 1, 2, 3, 4, 5, 6
        .EditCell
    End Select
    End With
    oShowGrid3
End If
If Index = 1 Then
    With ogrid(1)
    Select Case .col
    Case 0
        .EditCell
        For irow = 1 To .Rows - 1
        If Not srow_poss = irow Then
            .TextMatrix(irow, 0) = 0
        End If
    Next
    End Select
    
    
    
    End With
    
    
End If
If Index = 2 Then
    With ogrid(2)
    Select Case .col
    Case 0
        .EditCell
    Case 3
        If .TextMatrix(.row, 0) = -1 Then
         .EditCell
        End If
    Case 4
        If .TextMatrix(.row, 0) = -1 Then
         .EditCell
        End If
'        For irow = 1 To .Rows - 1
'        If Not srow_poss = irow Then
'            .TextMatrix(irow, 0) = 0
'        End If
'    Next
    End Select
    
    
    
    End With
    
    
End If
End Sub

Private Sub searchby_Click()
oShowGrid1
End Sub

Private Sub searchby_KeyDown(KeyCode As Integer, Shift As Integer)
oShowGrid1
End Sub

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "keyCustomer"
            Frame1(6).ZOrder   'Picture1(0).ZOrder
Case "keyProduct"
            Frame1(7).ZOrder   'Picture1(0).ZOrder
            Text1(18).SetFocus
Case "keyCompany"
            Frame1(8).ZOrder
Case "keyInsentif"
            Frame1(9).ZOrder
End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "Informasi Produk"
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 0, 8, 9, 18
    If Check2.value = Checked Then
        oShowGrid1
    End If
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

Public Sub oShowGrid1()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "call sp_master_produk_harga_customer_get('"
    sQuery = sQuery & Text1(0) & "','" & Text1(9) & "','"
    sQuery = sQuery & searchby.ListIndex + 1 & "','%" & Text1(18) & "%')"


    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(0)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid(0)
       .Cols = .Cols + 1
        .ColHidden(.Cols - 1) = True
         .ColHidden(.Cols - 2) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodeproduk"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namaproduk_cust"))
                .TextMatrix(i, 2) = ToNumber(oRsDetail("harga"))
                .TextMatrix(i, 3) = ToNumber(oRsDetail("fee"))
                .TextMatrix(i, 4) = ToNumber(oRsDetail("harga_sales"))
                .TextMatrix(i, 5) = ToNumber(oRsDetail("harga_purchase"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("kodeproduk_custmr"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("namaproduk"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select lasposrow, 0
            Label1(21).Caption = .TextMatrix(.row, 0)
            Label1(22).Caption = .TextMatrix(.row, 5)
            Label1(23).Caption = .TextMatrix(.row, 4)
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub SimpanGrid1(keykodecustomer As String, keyKodeHarga As String)
On Error GoTo errhandler

    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With ogrid(0)
        For irow = 1 To .Rows - 1
            sQuery = "call sp_master_produk_harga_customer_update_new('" & keykodecustomer & "','"
            sQuery = sQuery & .TextMatrix(irow, 0) & "','"
            sQuery = sQuery & keyKodeHarga & "','"
            sQuery = sQuery & RTrim(.TextMatrix(irow, 1)) & "','"
            sQuery = sQuery & toNumberIndonesia(.TextMatrix(irow, 2)) & "','"
            sQuery = sQuery & toNumberIndonesia(.TextMatrix(irow, 3)) & "','"
            sQuery = sQuery & toNumberIndonesia(.TextMatrix(irow, 4)) & "','"
            sQuery = sQuery & toNumberIndonesia(.TextMatrix(irow, 5)) & "','"
            sQuery = sQuery & MenuFrm.sUserID & "')"
            If .TextMatrix(irow, .Cols - 1) = "2" Then
                oKon.Execute (sQuery)
            End If
        Next
        
        
        End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub
Public Sub oShowGrid2()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "call sp_master_customer_company_get('"
    sQuery = sQuery & Text1(0) & "')"


    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(1)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid(1)
       .Cols = .Cols + 1
        .ColHidden(.Cols - 1) = True
        
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = IIf(RTrim(oRsDetail("aktive")) = "Y", -1, 0)
                .TextMatrix(i, 1) = RTrim(oRsDetail("CmpnyID"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("CmnyName"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("address"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("City"))
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

Public Sub oShowGrid3()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "call sp_master_produk_harga_customer_insentif_entry_get('"
    sQuery = sQuery & Text1(0) & "','"
    sQuery = sQuery & Label1(21).Caption & "','"
    sQuery = sQuery & Text1(9) & "')"


    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(2)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid(2)
       .Cols = .Cols + 1
        .ColHidden(.Cols - 1) = True
        
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = IIf(RTrim(oRsDetail("pilih")) = "Y", -1, 0)
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodesalesman"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("namasalesman"))
                .TextMatrix(i, 3) = IIf(RTrim(oRsDetail("isinsentif")) = "1", -1, 0)
                .TextMatrix(i, 4) = ToNumber(oRsDetail("insentif"))
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


Public Sub SimpanGrid3()
On Error GoTo errhandler


    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With ogrid(2)
        For irow = 1 To .Rows - 1
            If .TextMatrix(irow, 0) = -1 Then
                sQuery = "call sp_master_produk_harga_customer_insentif_insert_update('"
                sQuery = sQuery & Text1(0) & "','"
                sQuery = sQuery & Label1(21).Caption & "','"
                sQuery = sQuery & Text1(9) & "','"
                sQuery = sQuery & .TextMatrix(irow, 1) & "','"
                sQuery = sQuery & .TextMatrix(irow, 4) & "','"
                sQuery = sQuery & IIf(.TextMatrix(irow, 0) = -1, "1", "0") & "','"
                sQuery = sQuery & MenuFrm.sUserID & "')"
                oKon.Execute (sQuery)
            End If
         Next
        
        End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub

Public Sub SimpanGrid2(keykodecustomer As String)
On Error GoTo errhandler


    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim is_update As Boolean
    is_update = True
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With ogrid(1)
        For irow = 1 To .Rows - 1
            sQuery = "call sp_master_customer_company_update('" & keykodecustomer & "','"
            sQuery = sQuery & .TextMatrix(irow, 1) & "','"
            sQuery = sQuery & IIf(.TextMatrix(irow, 0) = -1, "Y", "N") & "','"
            sQuery = sQuery & MenuFrm.sUserID & "')"
            oKon.Execute (sQuery)
         Next
        
        End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form LhppFromTF 
   BackColor       =   &H8000000A&
   Caption         =   "Penerbitan Lembar Tukar Faktur Form"
   ClientHeight    =   8100
   ClientLeft      =   -135
   ClientTop       =   975
   ClientWidth     =   12195
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
   Icon            =   "LhppFromTF.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   38
      Top             =   1440
      Width           =   12495
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   420
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   741
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   13
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   9
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   8
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   10
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   19
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1455
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   0
      Width           =   12495
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Index           =   11
         Left            =   1920
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   600
         Width           =   1650
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Index           =   10
         Left            =   1920
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   960
         Width           =   1650
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
         Left            =   9720
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Auto"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   8880
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Close"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   10800
         TabIndex        =   21
         Top             =   960
         Width           =   855
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
         Left            =   1920
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   240
         Width           =   1650
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
         Left            =   4140
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   3315
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   19
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LhppFromTF.frx":C84A
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
         Left            =   9720
         TabIndex        =   23
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
         Left            =   11760
         TabIndex        =   26
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LhppFromTF.frx":C866
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
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Open"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   22
         Top             =   960
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Entry"
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Total Nilai LHPP"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "No.Penerimaan"
         Height          =   315
         Index           =   0
         Left            =   7560
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   2
         Left            =   7560
         TabIndex        =   28
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Status Dokumen"
         Height          =   315
         Index           =   4
         Left            =   7560
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Kolektor"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1095
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   7080
      Width           =   12495
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
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   5055
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
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Referensi"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   3615
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   12495
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   12135
         _cx             =   21405
         _cy             =   5741
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"LhppFromTF.frx":C882
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
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   12495
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Index           =   7
         Left            =   9720
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Index           =   6
         Left            =   9720
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   240
         Width           =   2535
      End
      Begin VSDFLATS.FlatButton FlatButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         MouseIcon       =   "LhppFromTF.frx":C9F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ambil Faktur"
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
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   600
         Width           =   5535
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
         Index           =   2
         Left            =   4140
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   3315
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
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1650
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "LhppFromTF.frx":CA11
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
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Total Nilai Kwitansi"
         Height          =   315
         Index           =   8
         Left            =   7560
         TabIndex        =   33
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Kwitansi"
         Height          =   315
         Index           =   7
         Left            =   7560
         TabIndex        =   32
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Customer"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16776960
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":CA2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":CD47
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D061
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D37B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D4D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D62F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D789
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":D8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":DA3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":DB97
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":DCF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":DE4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":DFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E0FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E259
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E3B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E50D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E667
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LhppFromTF.frx":E7C1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbllinenum 
      Caption         =   "1"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   10320
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "LhppFromTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim terbilang As New CRUFLFungsiku.Konversi
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim iNewEntrySts As Integer
Dim iNilLHHPnotCurrentEntry As Double

Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String

Dim sUpdateEntry As String
Dim sInsertEntry As String
Dim sDeleteEntry As String

Dim istatus As StatusForm

Dim sbatchid As Integer
Dim snodokumen As String
Dim stgldokumen As Date
Dim sdokstatus As String
Dim skodekolektor As String
Dim sketerangan As String
Dim sreferensi As String
Dim sjmlentry As Integer
Dim sjmlkwitansi As Integer
Dim sjmllhpp As Integer
Dim stotnillhpp As Double
Dim sobjtype As Integer
Dim saudituser As String
Dim sauditdate As Date
Dim sdocentry As Double
Dim sdocentry_sts As String
Dim skodecustomer As String
Dim stotnilkwitansi As Double

Dim slinenum As Integer
Dim snodokumen2 As String
Dim stgldokumen2 As Date
Dim sdokstatus2 As String
Dim skodecustomer2 As String
Dim stotalsetppn As Double
Dim ssisanilkwitansi As Double

Dim sbasedocentry As Double
Dim slinenummax As Integer



Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = " call sp_transaksi_lhpp_tf_get('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
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
    sQuery = "call sp_transaksi_lhpp_tf_get('" & Text1(0).text & "',1)"
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
    sQuery = "call sp_transaksi_lhpp_tf_get('" & Text1(0).text & "',3)"
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
    sQuery = "call sp_transaksi_lhpp_tf_get('" & Text1(0).text & "',2)"
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
    sQuery = "call sp_transaksi_lhpp_tf_get('" & Text1(0).text & "',4)"
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
        MsgBox "Kode Kolektor Belum Di Entri ", vbInformation
        Text1(9).SetFocus
        Exit Sub
    End If
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             FindData Text1(0)
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
             FlatButton1.Enabled = True
        End If
    End If
End Sub
Public Sub DeleteData()
    Dim ires As Integer
    
    If oFindByQuery("SELECT  tukarfaktur FROM transaksi_keluar_tfaktur WHERE nodokumen='" & Text1(0) & "'", DBaseConection.Modul) = "Y" Then
        MsgBox "Dokumen " & Text1(0) & " Tidak Bisa di Hapus , Sudah Dilakukan Tukar Faktur", vbInformation, "Data Tukar Faktur"
        FindData Text1(0)
        Exit Sub
    End If
    
    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
    If ires = 6 Then
        If DoDeleteData Then
             MsgBox "Data Sudah Terhapus", , "Hapus Data"

             MovePrevious

             
        End If
    End If
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = StatusForm.DataBaru Then
            sQuery = sInsert
            Set oRs = oCon.Execute(sQuery)
            sbatchid = oRs(0)
            
        Else
            sQuery = sUpdate
            oCon.Execute sQuery
        End If
        
        SaveEntry sbatchid, sdocentry, 1

        'oCon.Execute "CALL sp_mecah_terbilang('" & terbilang.terbilang(ToNumber(Text1(6))) & "',55," & sdocentry & ")"
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
        sQuery = sDelete ' "Delete from transaksi_keluar where nodokumen='" & snodokumen & "'"
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
        If oCekJumlahTrx("transaksi_keluar", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    FlatDatePicker1.value = Now()
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
    Text1(0).Locked = False
    'Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    'Option2(0).value = True
    GridModul.ClearGridDetail ogrid
    lbllinenum.Caption = 1
    Option1(0).value = True
    iNewEntrySts = 1
    Text1(11) = 0
    Text1(10) = 0
    
    oToolBarsEnable 1, False
    clearnewentry
    BrowseUserID(1).Enabled = False
    FlatButton1.Enabled = False
   
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    snodokumen = ToText(Text1(0).text)
    'sdocentry = oRs("docentry")
    If istatus = StatusForm.DataBaru Then
        snodokumen = IIf(Text1(0).text = "", GetDocnum(transaksi_lhpp_tf, True, DBaseConection.Modul), Text1(0).text)
        Text1(0).text = snodokumen
    Else
        snodokumen = ToText(Text1(0).text)
    End If
    skodecustomer = Text1(1).text
    stgldokumen = FlatDatePicker1.value
    If Option1(0).value = True Then
        sdokstatus = "1"
    Else
        sdokstatus = "0"
    End If
    skodekolektor = Text1(9)
    sketerangan = Text1(4)
    sreferensi = Text1(5)
    sjmlentry = ToNumber(Text1(11))
    sjmlkwitansi = ToNumber(Text1(11))
    stotnillhpp = toNumberIndonesia(ToNumber(Text1(10)))

    sQuery = "call sp_transaksi_lhpp_tf_update('" & sbatchid & "','"
    sQuery = sQuery & snodokumen & "','"
    sQuery = sQuery & Format(stgldokumen, "YYYY-MM-DD") & "','"
    sQuery = sQuery & sdokstatus & "','"
    sQuery = sQuery & skodekolektor & "','"
    sQuery = sQuery & sketerangan & "','"
    sQuery = sQuery & sreferensi & "','"
    sQuery = sQuery & sjmlentry & "','"
    sQuery = sQuery & sjmlkwitansi & "','"
    sQuery = sQuery & stotnillhpp & "','"
    sQuery = sQuery & saudituser & "')"

    sUpdate = sQuery
    sInsert = Replace(sQuery, "update", "insert")
    sDelete = Replace(sQuery, "update", "delete")

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
oBrowse.ShowFinder Browslhpptf, "", ubDescending, DBaseConection.Modul
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
        Showcustomer Text1(1)
        
    End If
Case 2
    oBrowse.ShowFinder BrowsKolektor, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(9) = oBrowse.YangDipilih
        Text1(8) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder Browslhpp, "dokstatus='1'", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(12) = oBrowse.YangDipilih
        'Text1(8) = oBrowse.Keterangan
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
         .TextMatrix(.row, 3) = oFindByQuery("select harga from master_produk_harga_customer where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "' and kodecustomer='" & Text1(1) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 4) = oFindByQuery("select fee from master_produk_harga_customer where kodecustomer='" & Text1(1) & "' and kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "' and kodecustomer='" & Text1(1) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 11) = oFindByQuery("select diskon from master_produk_diskon where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodediskon='" & Text1(5) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 5) = ToNumber(.TextMatrix(.row, 3)) * (ToNumber(.TextMatrix(.row, 11)) / 100)
       .Select .row, 2
        End With
    End If
Case 6
    oBrowse.ShowFinder BrowsSalesman, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(17) = oBrowse.YangDipilih
        Text1(18) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub


Private Sub Check1_Click(Index As Integer)
If Check1(0).value = 0 Then
    Text1(0).Enabled = True
Else
    Text1(0).Enabled = False
End If
End Sub

Private Sub FlatButton1_Click()
If Text1(1) = "" Then
    MsgBox "Customer Kosong , Pilih Customer ", vbInformation
    Exit Sub
End If
Dim oBrowseDaftarlhpp As New MonitoringKwitansiBrowseTF
oBrowseDaftarlhpp.ShowForm ogrid, 1, 1, Text1(1)
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Penerbitan Lembar Tukar Faktur"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
iNewEntrySts = 0
oFormatOption 1, Me
oFormatCheckList 1, Me
cleardata
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    sbatchid = oRs("batchid")
    sdocentry = IIf(IsNull(oRs("docentry")), 0, oRs("docentry"))
    Text1(0).text = oRs("nodokumen")
    KodeUserAksesTemp = oRs("nodokumen")
    Text1(0).Locked = True
    'Text1(1).text = oRs("kodekolektor")
    FlatDatePicker1.value = oRs("tgldokumen")
    
    Text1(9) = oRs("kodekolektor")
    Text1(8) = oRs("namakolektor")
    
    Text1(4) = oRs("keterangan")
    Text1(5) = oRs("referensi")
    
    Text1(11) = (oRs("jmlentry"))
    Text1(10) = formatRupiah(oRs("totnillhpp"))
    oShowEntry sbatchid, sdocentry, 0
    oToolBarsEnable 0, True
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    'istatus = Normal
    If "1" = oRs("dokstatus") Then
    Option1(0).value = True
    Option1(1).value = False
        istatus = StatusForm.NormalPlusExec
    Else
        istatus = StatusForm.NormalClosePlusExec
    Option1(0).value = False
    Option1(1).value = True
    End If
    
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
    FlatButton1.Enabled = True
    iNewEntrySts = 0
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
clearnewentry
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub



Private Sub ogrid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'BrowseUserID(5).Visible = False
End Sub

Private Sub ogrid_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = ogrid.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, ogrid, istatus
With ogrid
Select Case col
Case 3
    'oRecalculate
End Select
End With
oRecalculate
End Sub

Private Sub ogrid_Click()
'With ogrid
'If .Rows = 1 And Not Text1(1) = "" Then
'    AddRow
'End If
'End With
End Sub

Private Sub ogrid_EnterCell()
'With ogrid
'    BrowseUserID(2).Visible = False
'    Select Case .col
'        Case 0
'            If .Rows = 1 Then Exit Sub
'                            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
'                            SetFinder BrowseUserID(5), ogrid, .col
'                            '.EditCell
'        Case 2, 3, 4, 6
'            .EditCell
'        Case 7
'
'        If .Rows - 1 = .row Then
'            AddRow
'            '.Select .row, 0
'        Else
'            .Select .row + 1, 0
'        End If
'
'    End Select
'End With
End Sub

Private Sub oGrid_GotFocus()
'With ogrid
'    'BrowseUserID(5).Visible = False
'    Select Case .col
'        Case 0
'            If .Rows = 1 Then Exit Sub
'                            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
'                            SetFinder BrowseUserID(5), ogrid, .col
'                            '.EditCell
'
'
'         Case 2
'            .EditCell
'
'
'    End Select
'End With
End Sub

Private Sub ogrid_KeyDown(KeyCode As Integer, Shift As Integer)
With ogrid

' If KeyCode = vbKeyDelete And Not .Rows = 1 Then
'        .TextMatrix(.row, 3) = 0
'        .TextMatrix(.row, 6) = 0
'        .TextMatrix(.row, 7) = 0
'        oRecalculate
' End If
    
MainModule.DoKeyDown KeyCode, istatus

    'If ToNumber(.TextMatrix(.Row, .Cols - 1)) = 0 Then Exit Sub
    If Not KeyCode = vbKeyInsert Then
           gridDetail_KeyDown KeyCode, 0, ogrid, istatus
           If KeyCode = vbKeyDelete Then Exit Sub
           Select Case .col
           Case 0
                If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                .EditCell
           Case 2, 3, 4, 6
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
With ogrid
If Not KeyCode = 13 Then Exit Sub

Select Case col
Case 0
    If .TextMatrix(.row, .Cols - 1) = "0" Then Exit Sub
    .Select .row, 0
    If oFindByQuery("select namaproduk from master_produk where kodeproduk='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul) = "" Then
        MsgBox "Master Produk Tidak Ditemukan", vbInformation
        .Select .row, 0
    Else
         .TextMatrix(.row, 1) = oFindByQuery("select namaproduk from master_produk where kodeproduk='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 3) = oFindByQuery("select harga from master_produk_harga_customer where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "' and kodecustomer='" & Text1(1) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 4) = oFindByQuery("select fee from master_produk_harga_customer where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodeharga='" & Text1(6) & "' and kodecustomer='" & Text1(1) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 11) = oFindByQuery("select diskon from master_produk_diskon where kodeproduk='" & .TextMatrix(.row, 0) & "' and kodediskon='" & Text1(5) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 6) = ToNumber(.TextMatrix(.row, 3)) * (ToNumber(.TextMatrix(.row, 11)) / 100)
       .Select .row, 2
       .EditCell
    End If

Case 2
    .Select .row, 3
Case 3
    .Select .row, 4
Case 4
    .Select .row, 6
Case 6
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
Case 7
Text1(10) = formatRupiah(ToNumber(Text1(7)) + iNilLHHPnotCurrentEntry)
'Case 15, 16
'    Text1(Index).Text = formatRupiah(ToNumber(Text1(Index).Text))
'    Text1(Index).SelStart = Len(Text1(Index).Text)
'    Text1(19) = Format(ToNumber(Text1(15)) * ToNumber(Text1(14)) / 100, "###,###,###.#0")
'    Text1(20) = Format(ToNumber(Text1(14)) + ToNumber(Text1(19)), "###,###,###.#0")
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

Public Sub ShowGrid(sbatchid As Integer, sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "CALL sp_transaksi_lhpp_tf_detail2_get('" & sbatchid & "','"
    sQuery = sQuery & sdocentry & "')"

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

        '.COLWIDTH(1) = .Width - (.COLWIDTH(0) + .COLWIDTH(2) + .COLWIDTH(3)) - 100
        GridModul.ClearGridDetail ogrid
        .ColHidden(.Cols - 1) = True
        .ColHidden(.Cols - 2) = True
        .ColHidden(.Cols - 3) = True
        .ColHidden(.Cols - 4) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("nodokumen"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("tgldokumen"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("kodecustomer"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("namacustomer"))
                .TextMatrix(i, 4) = ToNumber(RTrim(oRsDetail("totalsetppn")))
                .TextMatrix(i, 5) = ToNumber(RTrim(oRsDetail("sisanilkwitansi")))
                .TextMatrix(i, .Cols - 4) = RTrim(oRsDetail("batchid"))
                .TextMatrix(i, .Cols - 3) = RTrim(oRsDetail("docentry"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("linenum"))
                

                .TextMatrix(i, .Cols - 1) = 0
                oRsDetail.MoveNext
            Loop
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
               lbllinenum.Caption = .TextMatrix(i, 8) + 1
        End If
    End With
    oKon.Close
    oRecalculate
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub Showcustomer(skodecustomer As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    sKondisi = " Where kodecustomer='" & skodecustomer & "' limit 1 "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call sp_master_customer_mov('"
    sQuery = sQuery & skodecustomer & "',0)"

    Set oRsDetail = oKon.Execute(sQuery)
    If Not oRsDetail.EOF Then
        Text1(3) = oRsDetail("alamat1")
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
Dim stotalsetppn2 As Double
With ogrid
    stotalsetppn2 = 0
    sttlawal = 0
    For irow = 1 To .Rows - 1
        stotalsetppn2 = stotalsetppn2 + ToNumber(.TextMatrix(irow, 5))
        sttlawal = sttlawal + 1
    Next

        Text1(6) = sttlawal
        Text1(7) = Format(stotalsetppn2, "###,###,###.#0")
        
End With
End Sub
Public Sub AddRow()
With ogrid
If .TextMatrix(.row, 0) = "" Then Exit Sub

    'If .row < .Rows - 1 And .TextMatrix(.row + 1, 0) = "" Then Exit Sub
'        If .row < .Rows - 1 Then
'           .Select .row + 1, 0
'           '.EditCell
'        Else
            
            .Rows = .Rows + 1
            .Select .Rows - 1, 0
            .Cell(flexcpFontBold, .row, 0, , .Cols - 1) = vbNormal
            '.EditCell
'        End If
        .TextMatrix(.row, 2) = 1
        .TextMatrix(.row, 3) = 0
        .TextMatrix(.row, 4) = 0
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
Public Sub SaveGrid(sbatchid As Integer, sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    Dim sInsertDetail As String
    Dim sUpdateDetail As String
    Dim sDeleteDetail As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

            Dim i As Double
            Dim sjumlahseb As Double
            For i = 1 To .Rows - 1
                
                slinenum = .TextMatrix(i, .Cols - 2)
                snodokumen2 = .TextMatrix(i, 0)
                stgldokumen2 = .TextMatrix(i, 1)
                sdokstatus2 = "1"
                skodecustomer2 = .TextMatrix(i, 2)
                stotalsetppn = .TextMatrix(i, 4)
                ssisanilkwitansi = .TextMatrix(i, 5)
   
                sQuery = "Call sp_transaksi_lhpp_tf_detail2_insert('" & sbatchid & "','"
                sQuery = sQuery & sdocentry & "','"
                sQuery = sQuery & slinenum & "','"
                sQuery = sQuery & snodokumen2 & "','"
                sQuery = sQuery & Format(stgldokumen2, "YYYY-MM-DD") & "','"
                sQuery = sQuery & sdokstatus2 & "','"
                sQuery = sQuery & skodecustomer2 & "','"
                sQuery = sQuery & toNumberIndonesia(formatRupiah(stotalsetppn)) & "','"
                sQuery = sQuery & toNumberIndonesia(formatRupiah(ssisanilkwitansi)) & "')"
                
                sInsertDetail = sQuery
                sUpdateDetail = Replace(sQuery, "insert", "update")
                sDeleteDetail = Replace(sQuery, "insert", "delete")
                                
                Select Case ToNumber(.TextMatrix(i, .Cols - 1))
                Case 1
                         oKon.Execute sInsertDetail
                        
                Case 2
                        oKon.Execute sUpdateDetail
                        
                Case 3
                        oKon.Execute sDeleteDetail
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
            
'                skodeproduk1 = .TextMatrix(i, 0)
'                sjumlah1 = ToNumber(.TextMatrix(i, 2))
'                sharga1 = ToNumber(.TextMatrix(i, 3))
'                stotalsebdiskon1 = ToNumber(.TextMatrix(i, 4))
'                sdiskontotal1 = ToNumber(.TextMatrix(i, 5))
'                stotalsetdiskon1 = ToNumber(.TextMatrix(i, 6))
'                sjumlahseb = ToNumber(.TextMatrix(i, 7))
'                skodegudang1 = .TextMatrix(i, 8)
'                skodediskon1 = .TextMatrix(i, 9)
'                skodeharga1 = .TextMatrix(i, 10)
'                sdiskonpersen1 = .TextMatrix(i, 11)
'                slinenum1 = ToNumber(.TextMatrix(i, .Cols - 2))
'
'                        oKon.Execute "update master_inventori set stock=stock+" & sjumlahseb & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
'                        oKon.Execute "delete from transaksi_keluar_detail1 where docentry='" & sdocentry & "' and linenum='" & slinenum1 & "'"

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


Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
'CALL sp_transaksi_lhpp_tf_print_insert_get('',15,'Admin');

'sQuery = "call sp_transaksi_lhpp_tf_get_form('"
'sQuery = sQuery & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "','"
'sQuery = sQuery & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "','"
'sQuery = sQuery & Text1(0) & "','"
'sQuery = sQuery & Text1(0) & "',"

sQuery = "SELECT count(*) "
sQuery = sQuery & " FROM transaksi_lhpp_tf a INNER JOIN transaksi_lhpp_tf_detail1 a1 ON a.batchid=a1.batchid"
sQuery = sQuery & " WHERE a.tgldokumen between '" & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "' and '" & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "'"
sQuery = sQuery & " and a.nodokumen between '" & Text1(0) & "' and '" & Text1(0) & "'"
'


If oFindByQuery(sQuery, DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If

sQuery = "SELECT upper(a.kodekolektor) as kodekolektor,a.nodokumen,a1.kodecustomer,a1.totnilkwitansi"
sQuery = sQuery & " FROM transaksi_lhpp_tf a INNER JOIN transaksi_lhpp_tf_detail1 a1 ON a.batchid=a1.batchid"
sQuery = sQuery & " WHERE a.tgldokumen between '" & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "' and '" & Format(FlatDatePicker1.value, "YYYY-MM-DD") & "'"
sQuery = sQuery & " and a.nodokumen between '" & Text1(0) & "' and '" & Text1(0) & "'"


With arlhppForm
   ' .lblHeaderTrx = "Form LHPP"
'    .lblCompany1 = MenuFrm.txtHeader(0)
'    .lblCompany2 = MenuFrm.txtHeader(1)
'    .lblCompany3 = MenuFrm.txtHeader(2)
    .Field5.text = "Tukar Faktur"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery ' & "0)"
    
'    .lblkode.Caption = "Kode Custmr"
'    .lblketerangan.Caption = "Nama Customer"
'    .txtkodeproduk.DataField = "custmrcode"
'    .txtproductname.DataField = "custmrname"

'    .PageSettings.Orientation = ddOPortrait
    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
'    .PageSettings.LeftMargin = MenuFrm.skiri
'    .PageSettings.RightMargin = MenuFrm.skanan
    .Show
    '.lblterbilang = terbilang.terbilang(100)
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    'new
    oToolBarsEnable 1, True
    clearnewentry
    BrowseUserID(1).Enabled = True
    iNewEntrySts = 1
Case 2
    'undo
    oToolBarsEnable 2, True
    oShowEntry sbatchid, sdocentry, 0
    BrowseUserID(1).Enabled = False
Case 3
    'save
    oToolBarsEnable 3, True
    BrowseUserID(1).Enabled = False
    SaveEntry sbatchid, sdocentry, 1
    MsgBox "Data Sudah Disimpan !!", vbOKOnly
Case 4
    'delete
    If MsgBox("Data Yakin Akan Dihapus !!", vbYesNo) = vbYes Then
        clearnewentry
        oToolBarsEnable 4, True
        SaveEntry sbatchid, sdocentry, 0
        oShowEntry sbatchid, sdocentry, 2
    Else
    End If
Case 5
    'first
    oShowEntry sbatchid, sdocentry, 1
Case 6
    'prev
    oShowEntry sbatchid, sdocentry, 2
Case 7
    'next
    oShowEntry sbatchid, sdocentry, 3
Case 8
    'last
    oShowEntry sbatchid, sdocentry, 4
Case 9
    
    Dim oBrowse As New BrowseFrm
    oBrowse.ShowFinder Browslhpptfdetail1, "batchid=" & sbatchid, ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then oShowEntry sbatchid, oBrowse.YangDipilih, 0
    Set oBrowse = Nothing

End Select
End Sub

Public Sub oShowEntry(sbatchidQ As Integer, sdocentryQ As Double, sget As Integer)
On Error GoTo errhandler
    Dim oKonQ As New ADODB.Connection
    Dim oRsDetailQ As New ADODB.Recordset
    Dim sKondisi As String
    sQuery = "sp_transaksi_lhpp_tf_detail1_get(" & sbatchidQ & "," & sdocentryQ & "," & sget & ")"
    iNilLHHPnotCurrentEntry = ToNumber(oFindByQuery("SELECT  SUM(sisanilkwitansi) FROM transaksi_lhpp_tf_detail2 WHERE batchid=" & sbatchid & " AND docentry!=" & sdocentry & "", DBaseConection.Modul))
    If oKonQ.State = 1 Then oKonQ.Close
    oKonQ.Open MainModule.Conectionku(DBaseConection.Modul)
    Set oRsDetailQ = oKonQ.Execute(sQuery)
    If Not oRsDetailQ.EOF Then
        sdocentry = oRsDetailQ("docentry")
        Text1(1) = oRsDetailQ("kodecustomer")
        Text1(2) = oRsDetailQ("namacustomer")
        Text1(3) = oRsDetailQ("alamat1")
'        istatus = Normal
'        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLhppFromTF
    End If
    oKonQ.Close
    ShowGrid sbatchid, sdocentry
    
    'Text1(3) = formatRupiah(iNilLHHPnotCurrentEntry)
    
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Delete Detail Data"
End Sub



Public Sub oToolBarsEnable(sget As Integer, istatus As Boolean)
Dim irows As Integer
With Toolbar1

If istatus = False Then

    For irows = 1 To .Buttons.Count
          .Buttons(irows).Enabled = False
    Next

Else
    For irows = 1 To .Buttons.Count
      .Buttons(irows).Enabled = True
    Next
    Select Case sget
    Case 0 'normal
        .Buttons(2).Enabled = False
    Case 1 'new
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = True
        .Buttons(4).Enabled = False
    Case 2 'undo
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(4).Enabled = True
    Case 3 'save
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(4).Enabled = True
    End Select

End If
End With
End Sub

Public Sub clearnewentry()
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(6) = "0"
Text1(7) = "0"
GridModul.ClearGridDetail ogrid
End Sub

Public Sub setEntryStatus()

    skodecustomer = Text1(1)
'    sjmlkwitansi = toNumberIndonesia(Text1(6))
'    stotnilkwitansi = toNumberIndonesia(Text1(7))
    sQuery = "call sp_transaksi_lhpp_tf_detail1_update('"
    sQuery = sQuery & sbatchid & "','"
    sQuery = sQuery & sdocentry & "','"
    sQuery = sQuery & skodecustomer & "','"
    sQuery = sQuery & toNumberIndonesia(Text1(6)) & "','"
    sQuery = sQuery & toNumberIndonesia(Text1(7)) & "')"
    
    sUpdateEntry = sQuery
    sInsertEntry = Replace(sQuery, "update", "insert")
    sDeleteEntry = Replace(sQuery, "update", "delete")
     
End Sub

Public Sub SaveEntry(sbatchid As Integer, sdocentry As Double, sentrytype As Integer)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    setEntryStatus
    If sentrytype = 1 Then
            If iNewEntrySts = 1 Then
               Set oRsDetail = oKon.Execute(sInsertEntry)
               sdocentry = oRsDetail(0)
            Else
               oKon.Execute sUpdateEntry
            End If
    Else
            oKon.Execute sDeleteEntry
    End If
    oKon.Close
            SaveGrid sbatchid, sdocentry
            ShowGrid sbatchid, sdocentry
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub


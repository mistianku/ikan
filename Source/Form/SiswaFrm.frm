VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form SiswaFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Anak Ke"
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
      Height          =   6255
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8535
      Begin VB.Frame Frame1 
         Caption         =   "Status"
         Height          =   1575
         Index           =   4
         Left            =   6840
         TabIndex        =   66
         Top             =   720
         Width           =   1575
         Begin VB.OptionButton Option3 
            Caption         =   "Tidak Aktif"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   69
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cuti"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Aktif"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   855
         End
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   64
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
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
         Left            =   2220
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   3600
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   3960
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   4320
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   51
         Text            =   "Text1"
         Top             =   4680
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   5040
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   3720
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   13
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   5760
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Tingkatan Sekolah"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   6735
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "DIATAS SMA"
            Height          =   255
            Index           =   4
            Left            =   3480
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "SMA"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   44
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "SMP"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   43
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "SD"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "TK"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Manual"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Laki Laki"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   26
         Top             =   2160
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Perempuan"
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   25
         Top             =   2160
         Width           =   2175
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   24
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaFrm.frx":0000
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
         Left            =   3120
         TabIndex        =   55
         Top             =   5760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaFrm.frx":001C
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
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   65
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
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
         Caption         =   "Kelas"
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   63
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Asal Sekolah"
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   62
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat "
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   61
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Telp. Rumah"
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   60
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Anak Ke"
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   59
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Dari"
         Height          =   315
         Index           =   11
         Left            =   3120
         TabIndex        =   58
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Bersaudara"
         Height          =   315
         Index           =   12
         Left            =   4680
         TabIndex        =   57
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Agama"
         Height          =   315
         Index           =   13
         Left            =   120
         TabIndex        =   56
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No Induk "
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No ID "
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Masuk"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Lengkap"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Panggilan"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tempat.Tanggal Lahir"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Jenis Kelamin"
         Height          =   315
         Index           =   22
         Left            =   120
         TabIndex        =   33
         Top             =   2160
         Width           =   2055
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8400
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "  Ayah  "
            Key             =   "keyAyah"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "    Ibu    "
            Key             =   "keyIbu"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Informasi Ayah"
      Height          =   1935
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   8535
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   16
         Left            =   3540
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1440
         Width           =   4575
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   17
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaFrm.frx":0038
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
         Caption         =   "Nama Ayah"
         Height          =   315
         Index           =   14
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Pekerjaan"
         Height          =   315
         Index           =   15
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No HP"
         Height          =   315
         Index           =   16
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Email"
         Height          =   315
         Index           =   17
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Informasi Ibu"
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   8535
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   21
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Left            =   2220
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   4575
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaFrm.frx":0054
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
         Caption         =   "Email"
         Height          =   315
         Index           =   18
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No HP"
         Height          =   315
         Index           =   19
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Pekerjaan"
         Height          =   315
         Index           =   20
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Ibu"
         Height          =   315
         Index           =   21
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "SiswaFrm"
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
Dim istatus As StatusForm

Dim stingkatansklh As String
Dim snoidsiswa As String
Dim sniksiswa As String
Dim stglmasuk As String
Dim snmlengkap As String
Dim snmpangilan As String
Dim sjnskelamin As String
Dim stptlahir As String
Dim stgllahir As String
Dim sKelas As String
Dim saslsekolah As String
Dim salmtrumah1 As String
Dim salmtrumah2 As String
Dim snotelprumah As String
Dim sanakke As Integer
Dim sdarike As Integer
Dim sagama As String
Dim snmayah As String
Dim spekerjaanayah As String
Dim snohpayah As String
Dim semailayah As String
Dim snamaibu As String
Dim spekerjaanibu As String
Dim snohpibu As String
Dim semailibu As String
Dim saudituser As String
Dim sauditdate As Date
Dim sstssiswa As String


Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from master_siswa where noidsiswa='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa
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
    sQuery = "Select *  from master_siswa order by noidsiswa asc limit 1"
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
    sQuery = "Select  *  from master_siswa where noidsiswa >'" & Text1(0).Text & "' order by noidsiswa asc limit 1"
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
    sQuery = "Select  *  from master_siswa where noidsiswa<'" & Text1(0).Text & "' order by noidsiswa desc limit 1"
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
    sQuery = "Select *  from master_siswa order by noidsiswa desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa
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
            If Option3(2).value = True Then
                oCon.Execute "update master_kelas set stskelas='0' where stskelas='1' and noidsiswa='" & snoidsiswa & "'"
            End If
            If Option3(2).value = False Then
                oCon.Execute "update master_kelas set stskelas='1' where stskelas='0' and noidsiswa='" & snoidsiswa & "'"
            End If
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
         oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "Delete from master_siswa where noidsiswa='" & snoidsiswa & "'"
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
        If oCekJumlahTrx("master_siswa", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    Check1(0).Enabled = True
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa
    If oFindByQuery("select autonodefault from setting_document where docid=" & master_siswa, parkir) = "0" Then
    Check1(0).value = 1
    Else
     Check1(0).value = 0
    End If
    
    'Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Check1(0).value = 0
    FlatDatePicker1(0).value = Format(Now(), "yyyy/mm/dd")
    FlatDatePicker1(1).value = Format(Now(), "yyyy/mm/dd")
    If Check1(0).value = 0 Then
        Text1(0).Enabled = False
        Text1(2).SetFocus
    End If
    
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    If istatus = StatusForm.DataBaru Then
        If Text1(0) = "" Then
            snoidsiswa = GetDocnum(master_siswa, True, parkir)
            Text1(0).Text = snoidsiswa
        Else
            snoidsiswa = ToText(Text1(0).Text)
        End If
    Else
        snoidsiswa = ToText(Text1(0).Text)
    End If
    
    'sniksiswa = GetDocnum(master_siswa, parkir)  'Text1(1).Text
    
    'snoidsiswa = Text1(0).Text
    sniksiswa = ToText(Text1(1).Text)
    
    stglmasuk = Format(FlatDatePicker1(0).value, "yyyy/mm/dd")
    snmlengkap = ToText(Text1(2).Text)
    snmpangilan = ToText(Text1(3).Text)
    sjnskelamin = IIf(Option1(0).value = True, "L", "P")
    stptlahir = ToText(Text1(4).Text)
    stgllahir = Format(FlatDatePicker1(1).value, "yyyy/mm/dd")
    sKelas = ToText(Text1(5).Text)
    saslsekolah = ToText(Text1(6).Text)
    salmtrumah1 = ToText(Text1(7).Text)
    salmtrumah2 = ToText(Text1(8).Text)
    snotelprumah = ToText(Text1(9).Text)
    sanakke = ToNumber(Text1(10).Text)
    sdarike = ToNumber(Text1(11).Text)
    sagama = Text1(12).Text
    snmayah = Replace(Text1(14).Text, "'", "\'")
    spekerjaanayah = Text1(15).Text
    snohpayah = Text1(17).Text
    semailayah = Text1(18).Text
    snamaibu = Replace(Text1(19).Text, "'", "\'")
    spekerjaanibu = Text1(20).Text
    snohpibu = Text1(22).Text
    semailibu = Text1(23).Text
    saudituser = MenuFrm.sUserID
    sauditdate = Now()
    sstssiswa = IIf(Option3(0).value = True, "1", IIf(Option3(1).value = True, "2", "0"))
    If Option2(0).value = True Then
        stingkatansklh = "0"
    End If
        If Option2(1).value = True Then
        stingkatansklh = "1"
    End If
        If Option2(2).value = True Then
        stingkatansklh = "2"
    End If
        If Option2(3).value = True Then
        stingkatansklh = "3"
    End If
        If Option2(4).value = True Then
        stingkatansklh = "4"
    End If
    
     
    sUpdate = "Update master_siswa set "
    sUpdate = sUpdate & "tingkatansklh='" & stingkatansklh & "',"
    sUpdate = sUpdate & "niksiswa='" & sniksiswa & "',"
    sUpdate = sUpdate & "tglmasuk='" & stglmasuk & "',"
    sUpdate = sUpdate & "nmlengkap='" & snmlengkap & "',"
    sUpdate = sUpdate & "nmpangilan='" & snmpangilan & "',"
    sUpdate = sUpdate & "jnskelamin='" & sjnskelamin & "',"
    sUpdate = sUpdate & "tptlahir='" & stptlahir & "',"
    sUpdate = sUpdate & "tgllahir='" & stgllahir & "',"
    sUpdate = sUpdate & "kelas='" & sKelas & "',"
    sUpdate = sUpdate & "aslsekolah='" & saslsekolah & "',"
    sUpdate = sUpdate & "almtrumah1='" & salmtrumah1 & "',"
    sUpdate = sUpdate & "almtrumah2='" & salmtrumah2 & "',"
    sUpdate = sUpdate & "notelprumah='" & snotelprumah & "',"
    sUpdate = sUpdate & "anakke='" & sanakke & "',"
    sUpdate = sUpdate & "darike='" & sdarike & "',"
    sUpdate = sUpdate & "agama='" & sagama & "',"
    sUpdate = sUpdate & "nmayah='" & snmayah & "',"
    sUpdate = sUpdate & "pekerjaanayah='" & spekerjaanayah & "',"
    sUpdate = sUpdate & "nohpayah='" & snohpayah & "',"
    sUpdate = sUpdate & "emailayah='" & semailayah & "',"
    sUpdate = sUpdate & "namaibu='" & snamaibu & "',"
    sUpdate = sUpdate & "pekerjaanibu='" & spekerjaanibu & "',"
    sUpdate = sUpdate & "nohpibu='" & snohpibu & "',"
    sUpdate = sUpdate & "emailibu='" & semailibu & "',"
    sUpdate = sUpdate & "audituser='" & saudituser & "',"
    sUpdate = sUpdate & "auditdate='" & sauditdate & "',"
    sUpdate = sUpdate & "stssiswa='" & sstssiswa & "'"
    sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "'"


    
    sInsert = "insert into master_siswa "
    sInsert = sInsert & " ("
    sInsert = sInsert & "noidsiswa,tingkatansklh,"
    sInsert = sInsert & "niksiswa,"
    sInsert = sInsert & "tglmasuk,"
    sInsert = sInsert & "nmlengkap,"
    sInsert = sInsert & "nmpangilan,"
    sInsert = sInsert & "jnskelamin,"
    sInsert = sInsert & "tptlahir,"
    sInsert = sInsert & "tgllahir,"
    sInsert = sInsert & "kelas,"
    sInsert = sInsert & "aslsekolah,"
    sInsert = sInsert & "almtrumah1,"
    sInsert = sInsert & "almtrumah2,"
    sInsert = sInsert & "notelprumah,"
    sInsert = sInsert & "anakke,"
    sInsert = sInsert & "darike,"
    sInsert = sInsert & "agama,"
    sInsert = sInsert & "nmayah,"
    sInsert = sInsert & "pekerjaanayah,"
    sInsert = sInsert & "nohpayah,"
    sInsert = sInsert & "emailayah,"
    sInsert = sInsert & "namaibu,"
    sInsert = sInsert & "pekerjaanibu,"
    sInsert = sInsert & "nohpibu,"
    sInsert = sInsert & "emailibu,"
    sInsert = sInsert & "audituser,stssiswa,"
    sInsert = sInsert & "auditdate)"
    sInsert = sInsert & " Values "
    sInsert = sInsert & "('" & snoidsiswa & "','"
    sInsert = sInsert & stingkatansklh & "','"
    sInsert = sInsert & sniksiswa & "','"
    sInsert = sInsert & stglmasuk & "','"
    sInsert = sInsert & snmlengkap & "','"
    sInsert = sInsert & snmpangilan & "','"
    sInsert = sInsert & sjnskelamin & "','"
    sInsert = sInsert & stptlahir & "','"
    sInsert = sInsert & stgllahir & "','"
    sInsert = sInsert & sKelas & "','"
    sInsert = sInsert & saslsekolah & "','"
    sInsert = sInsert & salmtrumah1 & "','"
    sInsert = sInsert & salmtrumah2 & "','"
    sInsert = sInsert & snotelprumah & "','"
    sInsert = sInsert & sanakke & "','"
    sInsert = sInsert & sdarike & "','"
    sInsert = sInsert & sagama & "','"
    sInsert = sInsert & snmayah & "','"
    sInsert = sInsert & spekerjaanayah & "','"
    sInsert = sInsert & snohpayah & "','"
    sInsert = sInsert & semailayah & "','"
    sInsert = sInsert & snamaibu & "','"
    sInsert = sInsert & spekerjaanibu & "','"
    sInsert = sInsert & snohpibu & "','"
    sInsert = sInsert & semailibu & "','"
    sInsert = sInsert & saudituser & "','"
    sInsert = sInsert & sstssiswa & "','"
    sInsert = sInsert & sauditdate & "')"
  
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
        Case 0
                oBrowse.ShowFinder BrowsSiswa, ""
                If Not oBrowse.YangDipilih = "" Then FindData (oBrowse.YangDipilih)
                Set oBrowse = Nothing
        Case 1
                oBrowse.ShowFinder BrowsAgama, ""
                If Not oBrowse.YangDipilih = "" Then
                    Text1(12).Text = oBrowse.YangDipilih
                    Text1(13).Text = oBrowse.Keterangan
                End If
                Set oBrowse = Nothing
        Case 2
                oBrowse.ShowFinder BrowsPekerjaan, ""
                If Not oBrowse.YangDipilih = "" Then
                    Text1(15).Text = oBrowse.YangDipilih
                    Text1(16).Text = oBrowse.Keterangan
                End If
                Set oBrowse = Nothing
        Case 3
                oBrowse.ShowFinder BrowsPekerjaan, ""
                If Not oBrowse.YangDipilih = "" Then
                    Text1(20).Text = oBrowse.YangDipilih
                    Text1(21).Text = oBrowse.Keterangan
                End If
                Set oBrowse = Nothing
End Select
        
End Sub



Private Sub Check1_Click(Index As Integer)
If Check1(0).value = 1 And istatus = DataBaru Then
    Text1(0).Enabled = True
    Text1(0).Locked = False
    Text1(0).SetFocus
Else
    Text1(0).Enabled = False
    Text1(0).Locked = True
End If
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterSiswa

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(12).Top
BrowseUserID(1).Height = Text1(12).Height
BrowseUserID(1).Left = Text1(12).Left + Text1(12).Width

BrowseUserID(2).Top = Text1(15).Top
BrowseUserID(2).Height = Text1(15).Height
BrowseUserID(2).Left = Text1(15).Left + Text1(15).Width

BrowseUserID(3).Top = Text1(20).Top
BrowseUserID(3).Height = Text1(20).Height
BrowseUserID(3).Left = Text1(20).Left + Text1(20).Width


End Sub

Private Sub Form_Load()

oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 3, Me
oFormatCheckList 1, Me
Frame1(1).ZOrder
istatus = Normal
cleardata
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Check1(0).Enabled = False
    Text1(0).Text = ToText(oRs("noidsiswa"))
    KodeUserAksesTemp = ToText(oRs("noidsiswa"))
    Text1(0).Locked = True
    Text1(1).Text = ToText(oRs("niksiswa"))
    FlatDatePicker1(0).value = oRs("tglmasuk")
    Text1(2).Text = ToText(oRs("nmlengkap"))
    Text1(3).Text = ToText(oRs("nmpangilan"))
    Option1(0) = IIf(ToText(oRs("jnskelamin")) = "L", True, False)
    Option1(1) = IIf(ToText(oRs("jnskelamin")) = "P", True, False)
    Text1(4).Text = ToText(oRs("tptlahir"))
    FlatDatePicker1(1).value = oRs("tgllahir")
    Text1(5).Text = ToText(oRs("kelas"))
    Text1(6).Text = ToText(oRs("aslsekolah"))
    Text1(7).Text = ToText(oRs("almtrumah1"))
    Text1(8).Text = ToText(oRs("almtrumah2"))
    Text1(9).Text = ToText(oRs("notelprumah"))
    Text1(10).Text = ToText(oRs("anakke"))
    Text1(11).Text = ToText(oRs("darike"))
    Text1(12).Text = ToText(oRs("agama"))
    Text1(13).Text = oFindByQuery("Select agama from master_agama where kode='" & ToText(oRs("agama")) & "'", parkir)
    Text1(14).Text = ToText(oRs("nmayah"))
    Text1(15).Text = ToText(oRs("pekerjaanayah"))
    Text1(16).Text = oFindByQuery("Select Pekerjaan from master_Pekerjaan where kode='" & ToText(oRs("pekerjaanayah")) & "'", parkir)
    Text1(17).Text = ToText(oRs("nohpayah"))
    Text1(18).Text = ToText(oRs("emailayah"))
    Text1(19).Text = ToText(oRs("namaibu"))
    Text1(20).Text = ToText(oRs("pekerjaanibu"))
    Text1(21).Text = oFindByQuery("Select Pekerjaan from master_Pekerjaan where kode='" & ToText(oRs("pekerjaanibu")) & "'", parkir)
    Text1(22).Text = ToText(oRs("nohpibu"))
    Text1(23).Text = ToText(oRs("emailibu"))
    Option2(ToNumber(oRs("tingkatansklh"))).value = True
    
    If ToNumber(oRs("stssiswa")) = "1" Then
        Option3(0).value = True
    End If
    If ToNumber(oRs("stssiswa")) = "2" Then
        Option3(1).value = True
    End If
    If ToNumber(oRs("stssiswa")) = "0" Then
        Option3(2).value = True
    End If
    
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

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "keyAyah"
            Frame1(1).ZOrder   'Picture1(0).ZOrder
Case "keyIbu"
            Frame1(0).ZOrder   'Picture1(0).ZOrder
End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "Informasi Produk"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000B
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
If Index = 0 And KeyCode = 13 Then FindData Text1(0).Text
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
'If Index = 0 Then FindData Text1(0).Text
End Sub

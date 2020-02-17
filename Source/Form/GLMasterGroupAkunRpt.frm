VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form GLMasterGroupAkunRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Group Entri Report"
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
      Caption         =   "Status"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "Semua Status"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tidak Aktif"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aktif"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode"
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   10455
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   10
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
         MouseIcon       =   "GLMasterGroupAkunRpt.frx":0000
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   11
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         MouseIcon       =   "GLMasterGroupAkunRpt.frx":001C
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox cr1 
      Height          =   480
      Left            =   1560
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   19
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kode Akun"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10455
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
         Left            =   4080
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   6135
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
         Left            =   4080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   6135
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
         Left            =   2220
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
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
         Index           =   0
         Left            =   2220
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLMasterGroupAkunRpt.frx":0038
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
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLMasterGroupAkunRpt.frx":0054
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
         Caption         =   "S/D Akun"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Akun"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GLMasterGroupAkunRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim syop As Integer
Dim smop As Integer
Dim skodefr As String
Dim skodeto As String
Dim sstatus As String
Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0

    oBrowse.ShowFinder BrowsAkunGroupSumberData, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsAkunGroupSumberData, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub

Private Sub Command1_Click()
Execution
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Group Entri Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLMasterGroupAkunRpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatFrameBackground Frame1(0)
oFormatOption 1, Me
istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"

Dim iTahun As Integer
For iTahun = 2011 To Year(Now()) + 5
FlatComboBox1(0).AddItem iTahun
Next
FlatComboBox1(0).text = Year(Now())
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Februari"
FlatComboBox1(1).AddItem "Maret"
FlatComboBox1(1).AddItem "April"
FlatComboBox1(1).AddItem "Mei"
FlatComboBox1(1).AddItem "Juni"
FlatComboBox1(1).AddItem "Juli"
FlatComboBox1(1).AddItem "Agustus"
FlatComboBox1(1).AddItem "September"
FlatComboBox1(1).AddItem "Oktober"
FlatComboBox1(1).AddItem "November"
FlatComboBox1(1).AddItem "Desember"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).ListIndex = Month(Now()) - 1

End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub
Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sKriteria As String

syop = FlatComboBox1(0).text
smop = FlatComboBox1(1).ListIndex + 1
skodefr = IIf(Text1(0) = "", oFindByQuery("select min(gr_dataentry) from tblgrupdataentry", DBaseConection.Modul), Text1(0))
skodeto = IIf(Text1(1) = "", oFindByQuery("select max(gr_dataentry) from tblgrupdataentry", DBaseConection.Modul), Text1(1))
sstatus = IIf(Option1(0).value = True, "Status='Y'", IIf(Option1(1).value = True, "Status='N'", "true"))
sKriteria = " where gr_dataentry between '" & skodefr & "'  and '" & skodeto & "' "
sQuery = "SELECT  gr_dataentry, nm_grupdata FROM  "
sQuery = sQuery & " tblgrupdataentry " & sKriteria
With arGLMasterGroupData
    .lblCompany1 = MenuFrm.txtHeader(0)
    .lblCompany2 = MenuFrm.txtHeader(1)
    .lblCompany3 = MenuFrm.txtHeader(2)
    .Label24.Caption = "Master Group Data Entri"
    .lblPeriode.Caption = "Kode Group Entri : " & skodefr & " s/d  " & skodeto
    .lblPeriode2.Visible = False
    '.lblPeriode2.Caption = "Status : " & IIf(Option1(0).value = True, "Aktif", IIf(Option1(1).value = True, "Tidak Aktif", "Semua"))
    
    
    '.lblPesan = stxtpesan
    .adoKu.Provider = "MSDASQL.1"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery
    
    .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar
    .Show
    If Not .adoKu.Recordset.EOF() Then
'    .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
'    .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value
    End If
End With




'Me.CR1.ParameterFields(0) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
'Me.CR1.ParameterFields(1) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
'Me.CR1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
'Me.CR1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
'Me.CR1.ParameterFields(4) = "yop" & ";" & syop & ";" & True
'Me.CR1.ParameterFields(5) = "mop" & ";" & sbulan & ";" & True

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


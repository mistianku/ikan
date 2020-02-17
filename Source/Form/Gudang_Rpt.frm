VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form Gudang_Rpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Gudang Report"
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
      Caption         =   "Gudang"
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1560
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
         Index           =   5
         Left            =   4080
         TabIndex        =   15
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
         Index           =   4
         Left            =   4080
         TabIndex        =   14
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
         Index           =   3
         Left            =   2220
         TabIndex        =   11
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
         Index           =   2
         Left            =   2220
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Gudang_Rpt.frx":0000
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
         Left            =   3600
         TabIndex        =   13
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Gudang_Rpt.frx":001C
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
         Caption         =   "S/D Kode Gudang"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari  Kode Gudang"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tipe Gudang"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   0
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
         Index           =   1
         Left            =   2220
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   7995
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "All Tipe Gudang"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Gudang_Rpt.frx":0038
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
         Caption         =   "Tipe Gudang"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Tipe Gudang"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Gudang_Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm

Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsTipeGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(4) = oBrowse.Keterangan
    End If
Case 2
    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(3) = oBrowse.YangDipilih
        Text1(5) = oBrowse.Keterangan
    End If
End Select

Set oBrowse = Nothing
End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    BrowseUserID(0).Enabled = False
Else
    Text1(0).Enabled = True
    Text1(1).Enabled = True
    BrowseUserID(0).Enabled = True
    Text1(0).SetFocus
End If
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Gudang Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudang_Rpt

End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me

oFormatFrameBackground Frame1(0)
oFormatFrameBackground Frame1(1)
If Check1.value = 1 Then
    Text1(0).Enabled = False
    Text1(1).Enabled = False
    BrowseUserID(0).Enabled = False
Else
    Text1(0).Enabled = True
    Text1(1).Enabled = True
    BrowseUserID(0).Enabled = True
    Text1(0).SetFocus
End If

istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

BrowseUserID(2).Top = Text1(3).Top
BrowseUserID(2).Height = Text1(3).Height
BrowseUserID(2).Left = Text1(3).Left + Text1(3).Width

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
    Text1(0).Enabled = False
    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler

Dim sKriteria As String
Dim stipegudang As String
Dim skodebrandfr As String
Dim skodebrandto As String
Dim txtmessage As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

skodebrandfr = IIf(Text1(2) = "", oFindByQuery("select kodegudang from master_gudang order by kodegudang asc limit 1", DBaseConection.Modul), Text1(2))
skodebrandto = IIf(Text1(3) = "", oFindByQuery("select kodegudang from master_gudang order by kodegudang desc limit 1", DBaseConection.Modul), Text1(3))


If Check1.value = 0 Then
    stipegudang = "tipegudang ='" & Text1(0) & "'"
Else
    stipegudang = "true"
End If
sKriteria = " where kodegudang  between '" & skodebrandfr & "' and '" & skodebrandto & "' and " & stipegudang

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vmaster_gudang vmaster_gudang1" & sKriteria
'
'
If oFindByQuery("select count(*) from vmaster_gudang vmaster_gudang1" & sKriteria, DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If
With arMasterGudang
    .lblCompany1 = MenuFrm.txtHeader(0)
    .lblCompany2 = MenuFrm.txtHeader(1)
    .lblCompany3 = MenuFrm.txtHeader(2)
    .Label24.Caption = "Master Harga Produk"
    .lblPeriode.Caption = "Kode Harga Produk : " & skodebrandfr & " s/d  " & skodebrandto
    .lblPeriode2.Visible = False

    .adoKu.Provider = "MSDASQL.1"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery
'    .lblKode = "Kode harga"
'    .lblKeterangan = "Keterangan"
'    .txtkode.DataField = "kodeharga"
'    .txtketerangan.DataField = "namaharga"
    .PageSettings.Orientation = ddOPortrait
'    .PageSettings.PaperHeight = MenuFrm.stinggi
'    .PageSettings.PaperWidth = MenuFrm.slebar

    .Show
    If Not .adoKu.Recordset.EOF() Then
'    .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
'    .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value

    End If
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


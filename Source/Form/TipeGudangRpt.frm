VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form TipeGudangRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Tipe Gudang Report"
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
      Caption         =   "Tipe Gudang"
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
         MouseIcon       =   "TipeGudangRpt.frx":0000
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
         MouseIcon       =   "TipeGudangRpt.frx":001C
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
         Caption         =   "S/D Tipe Gudang"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Tipe Gudang"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "TipeGudangRpt"
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

    oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Tipe Gudang Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTipeGudangRpt

End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"
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
Dim skodebrandfr As String
Dim skodebrandto As String

skodebrandfr = IIf(Text1(0) = "", oFindByQuery("select tipegudang from master_tipegudang order by tipegudang asc limit 1", (DBaseConection.Modul)), Text1(0))
skodebrandto = IIf(Text1(1) = "", oFindByQuery("select tipegudang from master_tipegudang order by tipegudang desc limit 1", DBaseConection.Modul), Text1(0))

Dim sKriteria As String

sKriteria = " where tipegudang  between '" & skodebrandfr & "' and '" & skodebrandto & "'"

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    master_tipegudang master_tipegudang1" & sKriteria
'
'

With arSatuanProduk
    .lblCompany1 = MenuFrm.txtHeader(0)
    .lblCompany2 = MenuFrm.txtHeader(1)
    .lblCompany3 = MenuFrm.txtHeader(2)
    .Label24.Caption = "Master Tipe Gudang"
    .lblPeriode.Caption = "Kode Tipe Gudang : " & skodebrandfr & " s/d  " & skodebrandto
    .lblPeriode2.Visible = False

    .adoKu.Provider = "MSDASQL.1"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery
    .lblKode = "Kode Tipe Gudang"
    .lblketerangan = "Keterangan"
    .txtkode.DataField = "tipegudang"
    .txtketerangan.DataField = "keterangan"
    .PageSettings.Orientation = ddOPortrait

    .Show
    If Not .adoKu.Recordset.EOF() Then

    End If
End With

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


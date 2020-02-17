VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form PerpindahanSiswaRpt 
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1560
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Peride Laporan"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
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
         MouseIcon       =   "PerpindahanSiswaRpt.frx":0000
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         MouseIcon       =   "PerpindahanSiswaRpt.frx":001C
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "PerpindahanSiswaRpt"
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

    oBrowse.ShowFinder BrowsBrand, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsBrand, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Laporan Perpindahan Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnPerpindahanSiswaRpt
End Sub

Private Sub Form_Load()
oFormatFrameBackground Frame1(0)
istatus = Normal
cleardata

oGetComboBoxTahun FlatComboBox1(0)
oGetComboBulanan FlatComboBox1(1)
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
    Text1(i).Text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim syop As Integer
Dim smop As Integer

syop = FlatComboBox1(0).Text
smop = FlatComboBox1(1).ListIndex + 1
'skodebrandfr = IIf(Text1(0) = "", oFindByQuery("select kodebrand from master_brand order by kodebrand asc limit 1", parkir), Text1(0))
'skodebrandto = IIf(Text1(1) = "", oFindByQuery("select kodebrand from master_brand order by kodebrand desc limit 1", parkir), Text1(0))

Me.CR1.Reset
Me.CR1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.CR1.ReportFileName = App.Path + "\Reports\perpindahansiswa_rpt.Rpt"

Dim sKriteria As String

sKriteria = " where yop=" & syop & " and mop=" & smop

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    master_perpindahan_siswa master_perpindahan_siswa1" & sKriteria
'
'
Me.CR1.SQLQuery = sQuery
Me.CR1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.CR1.ParameterFields(1) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.CR1.ParameterFields(2) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.CR1.ParameterFields(3) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True

Me.CR1.Destination = crptToWindow
Me.CR1.RetrieveDataFiles
Me.CR1.WindowState = crptMaximized
Me.CR1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


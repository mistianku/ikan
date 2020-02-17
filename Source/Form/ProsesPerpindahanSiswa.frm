VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form ProsesPerpindahanSiswa 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Akumulasi Perpindahan Siswa"
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   10455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Proses Akumulasi Perpindahan Siswa"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Border          =   0   'False
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
         MouseIcon       =   "ProsesPerpindahanSiswa.frx":0000
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   13
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         Border          =   0   'False
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
         MouseIcon       =   "ProsesPerpindahanSiswa.frx":001C
      End
      Begin VB.Label Label1 
         Caption         =   "Periode"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kode Brand"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
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
         MouseIcon       =   "ProsesPerpindahanSiswa.frx":0038
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
         MouseIcon       =   "ProsesPerpindahanSiswa.frx":0054
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
         Caption         =   "S/D Brand"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Brand"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "ProsesPerpindahanSiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim sproses1 As Integer
Dim sproses2 As Integer
Dim sproses3 As Integer




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
sTitle = "Proses Perpindahan Siswa By Periode"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesBulanan
End Sub

Private Sub Form_Load()
oFormatFrameBackground Frame1(0)
'oFormatOption 1, Me
istatus = Normal
cleardata
oAddComboBoxTahun
oGetComboBulanan
oFormatCheckList 1, Me
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
    Text1(i).Text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sTambah As Integer

If Check1(0).value = 1 Then
    
Else
    MsgBox "Centang Proses Akumulasi Perpindahan Siswa ", vbInformation
    Exit Sub
End If

sQuery = "Call sp_proses_master_perpindahan_siswa(" & FlatComboBox1(0).Text & "," & FlatComboBox1(1).ListIndex + 1 & ",'"
sQuery = sQuery & MenuFrm.sUserID & "')"

If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        oCon.Execute sQuery
oCon.Close
MsgBox "Proses Bulanan Selesai ", vbInformation
Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Proses Bulanan"
End Sub


Public Sub oAddComboBoxTahun()
Dim i As Integer
Dim sawal As Integer
sawal = ToNumber(oFindByQuery("select min(year(tglpendaftaran)) from transaksi_pendaftaran limit 1", parkir))
For i = IIf(sawal > Year(Now), Year(Now), sawal) To Year(Now)
    FlatComboBox1(0).AddItem i
Next
FlatComboBox1(0).Text = Year(Now())
End Sub


Public Sub oGetComboBulanan()
On Error GoTo errhandler
Dim i As Integer

        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "select namabulan from master_bulan"
        Set oRs = oCon.Execute(sQuery) 'master_moduleaccess
        With oRs
        Do While Not .EOF
            FlatComboBox1(1).AddItem .Fields(0)
        .MoveNext
        Loop
        FlatComboBox1(1).ListIndex = Month(Date) - 1
        End With
        oCon.Close
        
        istatus = Normal
        Exit Sub

errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Sub

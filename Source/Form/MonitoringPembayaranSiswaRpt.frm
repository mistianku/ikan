VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form MonitoringPembayaranSiswaRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   5835
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   11310
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
   ScaleWidth      =   11310
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Ditampilkan Secara"
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   10455
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rekap"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Rinci"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status Pembayaran"
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   10455
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cuti"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sudah Bayar"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Belum Bayar"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4200
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
      Caption         =   "Peride Pembayaran"
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
         MouseIcon       =   "MonitoringPembayaranSiswaRpt.frx":0000
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
         MouseIcon       =   "MonitoringPembayaranSiswaRpt.frx":001C
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
Attribute VB_Name = "MonitoringPembayaranSiswaRpt"
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
sTitle = "Laporan Monitoring Pembayaran Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringPembayaranSiswaRpt
End Sub

Private Sub Form_Load()
oFormatFrameBackground Frame1(0)
oFormatOption 2, Me
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

End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim syop As Integer
Dim smop As String
Dim sjnsbayar As String
Dim sdetail As String

If Option1(0).value = True Then
    sjnsbayar = "0"
End If
If Option1(1).value = True Then
    sjnsbayar = "3"
End If
If Option1(2).value = True Then
    sjnsbayar = "2"
End If


If Option2(0).value = True Then
    sdetail = "Y"
Else
    sdetail = "N"
End If

syop = (FlatComboBox1(0).Text)
smop = IIf(Len(Trim(FlatComboBox1(1).ListIndex + 1)) = 1, "0" & FlatComboBox1(1).ListIndex + 1, FlatComboBox1(1).ListIndex + 1)

Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\monitoring_pembayaran_rpt.rpt"

Dim sKriteria As String

sKriteria = " Where yop=" & syop & " and mop=" & smop & " and "
sKriteria = sKriteria & " jnsbayar ='" & sjnsbayar & "'"
sQuery = "Select * from vget_monitoring_pembayaran_rpt  vget_monitoring_pembayaran_rpt1 "
sQuery = sQuery & sKriteria

Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(1) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
Me.cr1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(4) = "yop" & ";" & syop & ";" & True
Me.cr1.ParameterFields(5) = "mop" & ";" & smop & ";" & True
Me.cr1.ParameterFields(6) = "jnsbayar" & ";" & sjnsbayar & ";" & True
Me.cr1.ParameterFields(7) = "detail" & ";" & sdetail & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form GLRab 
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
      Height          =   1935
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   14
         Top             =   240
         Width           =   1815
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1320
         Width           =   5295
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
         TabIndex        =   10
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
         Index           =   3
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   960
         Width           =   5295
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
         TabIndex        =   6
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
         Index           =   1
         Left            =   2340
         TabIndex        =   3
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
         Index           =   0
         Left            =   2340
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLRab.frx":0000
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
         Left            =   3720
         TabIndex        =   12
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLRab.frx":001C
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
         Left            =   3720
         TabIndex        =   13
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLRab.frx":0038
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
         Caption         =   "COA Kredit"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "COA Debet"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama RAB"
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode RAB"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GLRab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim skd_rab   As String
Dim snm_rab   As String
Dim srab_debet    As String
Dim srab_kredit   As String
Dim saktif As String
Dim KataKunci As String

Dim KodeUserAksesTemp As String
Dim sKodeUserAkses As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "spget_tblrab_mov('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
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
    sQuery = "spget_tblrab_mov('" & sKodeUserAkses & "',1)"
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
    sQuery = "spget_tblrab_mov('" & sKodeUserAkses & "',2)"
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
    sQuery = "spget_tblrab_mov('" & sKodeUserAkses & "',3)"
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
    sQuery = "spget_tblrab_mov('" & sKodeUserAkses & "',4)"
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skd_rab = Text1(0).Text
    snm_rab = Text1(1).Text
    srab_debet = Text1(2).Text
    srab_kredit = Text1(4).Text
    If Check1(0).value = 1 Then
        saktif = "Y"
    Else
        saktif = "N"
    End If
    
    sQuery = "('" & skd_rab & "' ,'"
    sQuery = sQuery & snm_rab & "' ,'"
    sQuery = sQuery & srab_debet & "' ,'"
    sQuery = sQuery & srab_kredit & "' ,'"
    sQuery = sQuery & saktif & "' ,'"
    sQuery = sQuery & MenuFrm.sUserID & "')"
    
    sUpdate = "spupdate_tblrab" & sQuery
    sInsert = "spinsert_tblrab" & sQuery
    sDelete = "spdelete_tblrab" & sQuery
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsAkunRAB, ""
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsAkunMasterCOA, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
Case 2
    oBrowse.ShowFinder BrowsAkunMasterCOA, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(4) = oBrowse.YangDipilih
        Text1(5) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsAkunGroupSumberData, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(8) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Rekening Antar Bagian"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMasterRAB
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

BrowseUserID(2).Top = Text1(4).Top
BrowseUserID(2).Height = Text1(4).Height
BrowseUserID(2).Left = Text1(4).Left + Text1(4).Width



End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatCheckList 1, Me
cleardata
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).Text = oRs("kd_rab")
    KodeUserAksesTemp = oRs("kd_rab")
    sKodeUserAkses = oRs("kd_rab")
    Text1(0).Locked = True
    Text1(1).Text = oRs("nm_rab")
    Text1(2).Text = oRs("rab_debet")
    Text1(3).Text = oRs("nm_akun_dbt")
    Text1(4).Text = oRs("rab_kredit")
    Text1(5).Text = oRs("nm_akun_krd")
    
    If oRs("aktif") = "Y" Then
        Check1(0).value = 1
    Else
        Check1(0).value = 0
    End If
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
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

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).Text
End Sub


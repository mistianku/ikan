VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form ModulFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Modul Aplikasi Form"
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
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
         Index           =   3
         Left            =   2340
         TabIndex        =   13
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
         Index           =   2
         Left            =   2340
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   960
         Width           =   7155
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Report"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   4680
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Transaksi"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Setting "
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   1680
         Value           =   -1  'True
         Width           =   1575
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
         Left            =   2340
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ModulFrm.frx":0000
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
         Caption         =   "Jenis Module"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Urutan Tampilan "
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Module"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Key Module"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Module ID"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "ModulFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sDscription As String
Dim svisorder As String
Dim stransid As String

Dim sModulID As String
Dim sModuleMenu As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select * from master_module where Modulid='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select *  from master_module order by Modulid asc limit 1"
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
    oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select  *  from master_module where Modulid >'" & Text1(0).text & "' order by Modulid asc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select  *  from master_module where Modulid<'" & Text1(0).text & "' order by Modulid desc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select *  from master_module order by Modulid desc limit 1 "
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
             MoveLast
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.login)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
        
        
        sQuery = "insert into master_moduleaccess "
        sQuery = sQuery & " SELECT kodegroup, modulid, IF(kodegroup='88',IF(transid=3,'N','Y'),'N') AS Baca,  "
        sQuery = sQuery & " IF(kodegroup='88',IF(transid=3,'N','Y'),'N') AS Tulis, "
        sQuery = sQuery & " IF(kodegroup='88',IF(transid=3,'N','Y'),'N') AS  Edit, "
        sQuery = sQuery & " IF(kodegroup='88',IF(transid=3,'N','Y'),'N') AS  Hapus, "
        sQuery = sQuery & " IF(kodegroup='88',IF(transid=3,'Y','Y'),'N') AS  Cetak"
        sQuery = sQuery & " FROM master_module ,master_group_user"
        sQuery = sQuery & " WHERE NOT EXISTS (SELECT * FROM master_moduleaccess WHERE Modulid=master_module.Modulid)"
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
         oCon.Open MainModule.Conectionku(DBaseConection.login)
        sQuery = "Delete from master_module where Modulid='" & sModulID & "'"
        oCon.Execute sQuery
        sQuery = "Delete from master_moduleaccess where Modulid='" & sModulID & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
    Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Text1(3) = oFindByQuery("select max(visorder) from master_module", DBaseConection.Modul) + 10
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    sModulID = ToText(Text1(0).text)
    sModuleMenu = ToText(Text1(1).text)
    sDscription = ToText(Text1(2).text)
    svisorder = ToText(Text1(3).text)
    stransid = IIf(Option1(0).value = True, "1", IIf(Option1(1).value = True, "2", "3"))
     
    sUpdate = "update master_module set "
    sUpdate = sUpdate & "ModuleMenu='" & sModuleMenu & "' , "
    sUpdate = sUpdate & "Dscription='" & sDscription & "',"
    sUpdate = sUpdate & "visorder='" & svisorder & "',"
    sUpdate = sUpdate & "transid='" & stransid & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "Modulid='" & sModulID & "'"

  
    sInsert = "insert into master_module ("
    sInsert = sInsert & "ModuleMenu,Dscription,visorder,transid ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & sModuleMenu & "',"
    sInsert = sInsert & "'" & sDscription & "',"
    sInsert = sInsert & "'" & svisorder & "',"
    sInsert = sInsert & "'" & stransid & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsModule, "", ubAscending, DBaseConection.login
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Modul Aplikasi"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnModulFrm
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 1, Me
cleardata
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("Modulid")
    KodeUserAksesTemp = oRs("Modulid")
    Text1(0).Locked = True
    Text1(1).text = oRs("ModuleMenu")
    Text1(2).text = oRs("Dscription")
    Text1(3).text = oRs("visorder")
    Select Case oRs("transid")
    Case "1"
            Option1(0).value = True
    Case "2"
            Option1(1).value = True
    Case "3"
            Option1(2).value = True
    End Select
    
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
    Text1(i).text = ""
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
If Index = 0 Then FindData Text1(0).text
End Sub

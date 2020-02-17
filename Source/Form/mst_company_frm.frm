VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form mst_company_frm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Company Form"
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
      BackColor       =   &H80000000&
      Height          =   1815
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
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1320
         Width           =   7515
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
         Width           =   7515
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
         Left            =   7980
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "mst_company_frm.frx":0000
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
         Caption         =   "Database Name"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Company"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Company"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dbase ID"
         Height          =   315
         Index           =   0
         Left            =   5880
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "mst_company_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sdbid As Integer
Dim scmpnyidQ As String
Dim scmpnyname As String
Dim sdatabasename As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "call sp_company_get('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        'MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, me.name,true
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
    sQuery = "call sp_company_get('" & KodeUserAksesTemp & "',1)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub


Public Sub MovePrevious()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "call sp_company_get('" & KodeUserAksesTemp & "',2)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MovePrevious"
End Sub
Public Sub MoveNext()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "call sp_company_get('" & KodeUserAksesTemp & "',3)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
    End If
    oCon.Close
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveNext"
End Sub
Public Sub MoveLast()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "call sp_company_get('" & KodeUserAksesTemp & "',4)"
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
    
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.login)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        oCon.Execute sQuery
        oCon.Close
        FindData Text1(1)
        Else
        sQuery = sUpdate
        oCon.Execute sQuery
        oCon.Close
        End If
        
        
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
        sQuery = sDelete
        oCon.Execute sQuery 'master_moduleaccess
                
        DoDeleteData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    KodeUserAksesTemp = Text1(1)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Text1(1).SetFocus
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    sdbid = ToNumber(Text1(0).text)
    scmpnyidQ = Text1(1).text
    scmpnyname = Text1(2).text
    sdatabasename = Text1(3).text
     
    sQuery = " Call sp_company_update('"
    sQuery = sQuery & sdbid & "','"
    sQuery = sQuery & scmpnyidQ & "','"
    sQuery = sQuery & scmpnyname & "','"
    sQuery = sQuery & sdatabasename & "','"
    sQuery = sQuery & MenuFrm.sUserID & "')"

    
    sUpdate = sQuery
    sInsert = Replace(sQuery, "update", "insert")
    sDelete = Replace(sQuery, "update", "delete")
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsCompany, "", ubAscending, DBaseConection.login
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Company"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
'MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnmst_company_frm
BrowseUserID.Top = Text1(1).Top
BrowseUserID.Height = Text1(1).Height
BrowseUserID.Left = Text1(1).Left + Text1(1).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
'oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
cleardata
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("dbid")
    KodeUserAksesTemp = oRs("cmpnyid")
    Text1(0).Locked = True
    Text1(1).text = oRs("cmpnyid")
    Text1(2).text = oRs("cmpnyname")
    Text1(3).text = oRs("databasename")
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

Private Sub oBorderStyle()
Dim irow
End Sub




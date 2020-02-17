VERSION 5.00
Begin VB.Form PelajaranFrm 
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
   Begin VB.CommandButton BrowseUserID 
      Caption         =   "..."
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   495
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
      Left            =   2580
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   735
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
      Left            =   2580
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   7155
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   8475
      Left            =   180
      Top             =   60
      Width           =   15
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   180
      Top             =   428
      Width           =   8835
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Pengguna Aplikasi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   300
      TabIndex        =   4
      Top             =   60
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "Kode"
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Agama"
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "PelajaranFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sKodeGroup As String
Dim snamagroup As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from master_pelajaran_group where kodegroup='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbar istatus
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
    sQuery = "Select *  from master_pelajaran_group order by kodegroup asc limit 1"
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
    sQuery = "Select  *  from master_pelajaran_group where kodegroup >'" & Text1(0).Text & "' order by kodegroup asc limit 1"
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
    sQuery = "Select  *  from master_pelajaran_group where kodegroup<'" & Text1(0).Text & "' order by kodegroup desc limit 1"
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
    sQuery = "Select *  from master_pelajaran_group order by kodegroup desc limit 1 "
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
             MenuFrm.SetToolbar istatus
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
    MenuFrm.SetToolbar istatus
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
        sQuery = "Delete from master_pelajaran_group where kodegroup='" & sKodeGroup & "'"
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
    MenuFrm.SetToolbar istatus
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbar istatus
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    sKodeGroup = Text1(0).Text
    snamagroup = Text1(1).Text
     
    sUpdate = "update master_pelajaran_group set "
    sUpdate = sUpdate & "namagroup='" & snamagroup & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "kodegroup='" & sKodeGroup & "'"
    
    sInsert = "insert into master_pelajaran_group ("
    sInsert = sInsert & "kodegroup,namagroup ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & sKodeGroup & "',"
    sInsert = sInsert & "'" & snamagroup & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsPelajaranGroup, ""
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Pelajaran"
lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbar istatus
MenuFrm.SetToolbarku istatus, MenuFrm.sGroupUserID, mnMateriGroup
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).Text = oRs("kodegroup")
    KodeUserAksesTemp = oRs("kodegroup")
    Text1(0).Locked = True
    Text1(1).Text = oRs("namagroup")
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
Text1(Index).BackColor = &H8000000B
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).Text
End Sub

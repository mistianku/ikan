VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form GudangFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Gudang Form"
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
      Height          =   3015
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   11775
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Tidak"
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Ya"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   22
         Top             =   2400
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
         Index           =   8
         Left            =   4200
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2040
         Width           =   5235
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
         Index           =   7
         Left            =   2340
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2040
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
         Index           =   6
         Left            =   2340
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1680
         Width           =   2955
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
         Left            =   2340
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1320
         Width           =   2955
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
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   960
         Width           =   2955
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
         Left            =   2340
         TabIndex        =   10
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
         Index           =   2
         Left            =   2340
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   7155
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   19
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MouseIcon       =   "GudangFrm.frx":0000
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
         Caption         =   "Cek Stok"
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tipe Gudang"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Faximale"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Telp"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kota"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   10440
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         MouseIcon       =   "GudangFrm.frx":001C
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
         Caption         =   "Nama Gudang"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Gudang"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GudangFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim skodegudang As String
Dim snamagudang As String
Dim salamat1 As String
Dim salamat2 As String
Dim skota As String
Dim stelp As String
Dim sFaximale As String
Dim stipegudang As String
Dim saktif As String
Dim scekstok As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from vmaster_gudang where kodegudang='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select *  from vmaster_gudang order by kodegudang asc limit 1"
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
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select  *  from vmaster_gudang where kodegudang >'" & Text1(0).text & "' order by kodegudang asc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select  *  from vmaster_gudang where kodegudang<'" & Text1(0).text & "' order by kodegudang desc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select *  from vmaster_gudang order by kodegudang desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
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
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "call sp_delete_master_gudang('" & skodegudang & "')"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Check1(0).value = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
Dim scall As String

    skodegudang = ToText(Text1(0).text)
    snamagudang = ToText(Text1(1).text)
     
    salamat1 = ToText(Text1(2))
    salamat2 = ToText(Text1(3))
    skota = ToText(Text1(4))
    stelp = ToText(Text1(5))
    sFaximale = ToText(Text1(6))
    stipegudang = ToText(Text1(7))
    If Option1(0).value = True Then
        scekstok = "1"
    Else
        scekstok = "0"
    End If
    If Check1(0).value = "1" Then
        saktif = "1"
    Else
        saktif = "0"
    End If
    
    scall = "('" & skodegudang & "','"
    scall = scall & snamagudang & "','"
    scall = scall & salamat1 & "','"
    scall = scall & salamat2 & "','"
    scall = scall & skota & "','"
    scall = scall & stelp & "','"
    scall = scall & sFaximale & "','"
    scall = scall & stipegudang & "','"
    scall = scall & saktif & "','"
    scall = scall & scekstok & "','"
    scall = scall & MenuFrm.sUserID & "','"
    scall = scall & Format(Now(), "YYYY-MM-DD") & "')"
    
    sUpdate = "call sp_update_master_gudang " & scall
    sInsert = "call sp_insert_master_gudang " & scall


'    sUpdate = "update master_gudang "
'    sUpdate = sUpdate & " set "
'    sUpdate = sUpdate & "namagudang= '" & snamagudang & "',"
'    sUpdate = sUpdate & "alamat1= '" & salamat1 & "',"
'    sUpdate = sUpdate & "alamat2= '" & salamat2 & "',"
'    sUpdate = sUpdate & "kota= '" & skota & "',"
'    sUpdate = sUpdate & "telp= '" & stelp & "',"
'    sUpdate = sUpdate & "faximale= '" & sFaximale & "',"
'    sUpdate = sUpdate & "tipegudang= '" & stipegudang & "',"
'    sUpdate = sUpdate & "aktif= '" & saktif & "',"
'    sUpdate = sUpdate & "cekstok= '" & scekstok & "',"
'    sUpdate = sUpdate & "audituser= '" & MenuFrm.sUserID & "',"
'    sUpdate = sUpdate & "auditdate= '" & Format(Now(), "YYYY-MM_DD") & "'"
'    sUpdate = sUpdate & " where "
'    sUpdate = sUpdate & "kodegudang='" & skodegudang & "'"
'
'    sInsert = "insert into master_gudang"
'    sInsert = sInsert & "("
'    sInsert = sInsert & "kodegudang,"
'    sInsert = sInsert & "namagudang,"
'    sInsert = sInsert & "alamat1,"
'    sInsert = sInsert & "alamat2,"
'    sInsert = sInsert & "kota,"
'    sInsert = sInsert & "telp,"
'    sInsert = sInsert & "faximale,"
'    sInsert = sInsert & "tipegudang,"
'    sInsert = sInsert & "aktif,cekstok,"
'    sInsert = sInsert & "audituser,"
'    sInsert = sInsert & "auditdate"
'    sInsert = sInsert & ")"
'    sInsert = sInsert & " values "
'    sInsert = sInsert & "("
'    sInsert = sInsert & "'" & skodegudang & "',"
'    sInsert = sInsert & "'" & snamagudang & "',"
'    sInsert = sInsert & "'" & salamat1 & "',"
'    sInsert = sInsert & "'" & salamat2 & "',"
'    sInsert = sInsert & "'" & skota & "',"
'    sInsert = sInsert & "'" & stelp & "',"
'    sInsert = sInsert & "'" & sFaximale & "',"
'    sInsert = sInsert & "'" & stipegudang & "',"
'    sInsert = sInsert & "'" & saktif & "',"
'    sInsert = sInsert & "'" & scekstok & "',"
'    sInsert = sInsert & "'" & MenuFrm.sUserID & "',"
'    sInsert = sInsert & "'" & Format(Now(), "YYYY-MM_DD") & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsTipeGudang, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(8) = oBrowse.Keterangan
    End If
End Select

Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Gudang"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGudangFrm

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(7).Top
BrowseUserID(1).Height = Text1(7).Height
BrowseUserID(1).Left = Text1(7).Left + Text1(7).Width
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me

oFormatOption 1, Me
oFormatCheckList 1, Me
cleardata
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("kodegudang")
    KodeUserAksesTemp = oRs("kodegudang")
    Text1(0).Locked = True
    Text1(1).text = oRs("namagudang")
    Text1(2) = oRs("alamat1")
    Text1(3) = oRs("alamat2")
    Text1(4) = oRs("kota")
    Text1(5) = oRs("telp")
    Text1(6) = oRs("faximale")
    Text1(7) = oRs("tipegudang")
    Text1(8) = oRs("namatipegudang")
    If oRs("cekstok") = "1" Then
       Option1(0).value = True
       Option1(1).value = False
    Else
       Option1(0).value = False
       Option1(1).value = True
    End If
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    
    If oRs("aktif") = "1" Then
        Check1(0).value = 1
    Else
        Check1(0).value = 0
    End If
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

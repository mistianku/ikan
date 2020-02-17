VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form UserFrm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Master Pengguna Form"
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
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5400
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Akses Unit"
            Key             =   "keyAksesUnit"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Akses Group Modul"
            Key             =   "keyAksesGroupModul"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10560
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Admin"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   8400
         TabIndex        =   9
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   8400
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   7
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "UserFrm.frx":0000
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
         ShowFocus       =   -1  'True
         ButtonBorderStyle=   4
         PictureOrientation=   2
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
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2220
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1320
         Width           =   8030
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
         Left            =   2220
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   8030
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
         Left            =   2220
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Password "
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nama User"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "User Id"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Akses Company"
      Height          =   3015
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   10575
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   2415
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   10335
         _cx             =   18230
         _cy             =   4260
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   255
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15790320
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"UserFrm.frx":001C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Akses Group Module"
      Height          =   3015
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   10575
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   2415
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   10335
         _cx             =   18230
         _cy             =   4260
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   255
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15790320
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"UserFrm.frx":00D4
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
End
Attribute VB_Name = "UserFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sKodeUserAkses As String
Dim kodeUserAkses As String
Dim namaUserAkses As String
Dim sadmin As String
Dim slocked As String
Dim KataKunci As String
Dim skodegroup As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "call sp_employee_get('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm
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
    sQuery = "call sp_employee_get('" & sKodeUserAkses & "',1)"
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
    sQuery = "call sp_employee_get('" & sKodeUserAkses & "',3)"
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
    sQuery = "call sp_employee_get('" & sKodeUserAkses & "',2)"
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
    sQuery = "call sp_employee_get('" & sKodeUserAkses & "',4)"
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
             FindData Text1(0)
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm
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
 
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm
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
        oCon.Close
        SaveGrid
        SaveGrid2
        DoSaveData = True
        If MenuFrm.isAdmin = "Y" Then
            istatus = Normal
        Else
            istatus = SettingForm
        End If
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm
    Text1(0).Locked = False
    Text1(0).Enabled = True
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Text1(2).TabIndex = 2
    ShowGrid Text1(0)
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    kodeUserAkses = Text1(0).text
    namaUserAkses = Text1(1).text
    KataKunci = EncryptPassword(Text1(2).text)
    
       
       slocked = IIf(Check1(0).value = 1, "N", "Y")
       sadmin = IIf(Check1(1).value = 1, "Y", "N")
      
    sQuery = "Call sp_employee_update('"
    sQuery = sQuery & kodeUserAkses & "','"
    sQuery = sQuery & namaUserAkses & "','"
    sQuery = sQuery & KataKunci & "','"
    sQuery = sQuery & slocked & "','"
    sQuery = sQuery & MenuFrm.sUserID & "','"
    sQuery = sQuery & sadmin & "')"
    sUpdate = "update master_User set "
    sUpdate = sQuery
    
    sInsert = Replace(sQuery, "update", "insert")
    sDelete = Replace(sQuery, "update", "delete")
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
oBrowse.ShowFinder BrowsUser, "", ubAscending, DBaseConection.login
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsUserGroup, "", ubAscending, login
    If Not oBrowse.YangDipilih = "" Then
        Text1(3) = oBrowse.YangDipilih
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Pengguna Form"
'MenuFrm.lblModul.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserFrm

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

'BrowseUserID(1).Top = Text1(3).Top
'BrowseUserID(1).Height = Text1(3).Height
'BrowseUserID(1).Left = Text1(3).Left + Text1(3).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatFrameBackground Frame1(0)
oFormatFrameBackground Frame2(0)
oFormatFrameBackground Frame2(1)
cleardata
Text1(0).Enabled = False
If MenuFrm.isAdmin = "Y" Then
            istatus = Normal
            BrowseUserID(0).Enabled = True
        Else
            istatus = SettingForm
            BrowseUserID(0).Enabled = False
        End If

If MenuFrm.sUserID = "admin" Then
    MoveFirst
Else
    Text1(1) = ""
    FindData MenuFrm.sUserID
    BrowseUserID(0).Enabled = False

End If
Frame2(0).ZOrder
FindData MenuFrm.sUserID
'Text1(1).SetFocus
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
'If LCase(oRs("UserID")) = "admin" Then
'    Text1(3).Enabled = False
'    Text1(4).Enabled = False
'    BrowseUserID(1).Enabled = True
'Else
'        Text1(3).Enabled = True
'        Text1(4).Enabled = True
'        BrowseUserID(1).Enabled = True
'End If

    Text1(0).text = oRs("emplcode")
    sKodeUserAkses = oRs("emplcode")
    KodeUserAksesTemp = oRs("emplcode")
    Text1(0).Locked = True
    Text1(1).text = oRs("emplname")
    Text1(2).text = DecryptPassword(oRs("pass"))
    Me.Caption = DecryptPassword(oRs("pass"))

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    ShowGrid Text1(0)
    ShowGrid2 Text1(0)
    If MenuFrm.isAdmin = "Y" Then
        Check1(0).Enabled = True
        Check1(1).Enabled = True
    Else
        Check1(0).Enabled = False
        Check1(1).Enabled = False
    End If

        Check1(1).value = IIf(oRs("admin") = "Y", 1, 0)
        Check1(0).value = IIf(oRs("locked") = "N", 1, 0)
    
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



Private Sub ogrid_CellChanged(Index As Integer, ByVal row As Long, ByVal col As Long)
Select Case Index
Case 0
    If row = 0 Then Exit Sub
    If col = ogrid(0).Cols - 1 Then Exit Sub
        GridModul.GridDetail_CellChanged row, col, ogrid(0), istatus
Case 1
    If row = 0 Then Exit Sub
    If col = ogrid(1).Cols - 1 Then Exit Sub
        GridModul.GridDetail_CellChanged row, col, ogrid(1), istatus
End Select
End Sub

Private Sub ogrid_Click(Index As Integer)
Select Case Index
Case 0
    With ogrid(0)
    Select Case .col
    Case 1
        If MenuFrm.isAdmin = "Y" Then
        .EditCell
        End If
    End Select
    End With
Case 1
    With ogrid(1)
    Select Case .col
    Case 1
        If MenuFrm.isAdmin = "Y" Then
        .EditCell
        End If
    End Select
    End With
End Select
End Sub

Private Sub TabStrip1_Click()
On Error GoTo errhandler
Select Case TabStrip1.SelectedItem.Key
Case "keyAksesUnit"
            Frame2(0).ZOrder   'Picture1(0).ZOrder
Case "keyAksesGroupModul"
            Frame2(1).ZOrder   'Picture1(0).ZOrder
End Select
Exit Sub
errhandler:
    MsgBox Err.Description, , "Informasi Master User"
End Sub

'Private Sub Text1_Change(Index As Integer)
'If Index = 3 Then
'    Text1(4) = oFindByQuery("select groupuser from master_group_user where kodegroup='" & Text1(3) & "'", modul)
'End If
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000B
End Sub

'Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'MainModule.DoKeyDown KeyCode, istatus
'End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).text
End Sub
Public Sub ShowGrid(sEmplCode As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    Dim sAkses As Integer
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "CALL sp_companyaccess_view('" & sEmplCode & "')"
    

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(0)

        GridModul.ClearGridDetail ogrid(0)
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                If oRsDetail("accessctrl") = "Y" Then
                    sAkses = -1
                Else
                    sAkses = 0
                End If
                
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("dbid"))
                .TextMatrix(i, 1) = sAkses
                .TextMatrix(i, 2) = RTrim(oRsDetail("cmpnyid"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("cmpnyname"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("dbid"))
                .TextMatrix(i, .Cols - 1) = 0
                oRsDetail.MoveNext
            Loop
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub ShowGrid2(sEmplCode As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    Dim sAkses As Integer
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "CALL sp_groupaccess_view('" & sEmplCode & "')"
    

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(1)

        GridModul.ClearGridDetail ogrid(1)
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                If oRsDetail("accessctrl") = "Y" Then
                    sAkses = -1
                Else
                    sAkses = 0
                End If
                
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("groupid"))
                .TextMatrix(i, 1) = sAkses
                .TextMatrix(i, 2) = RTrim(oRsDetail("kodegroup"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("groupuser"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("groupid"))
                .TextMatrix(i, .Cols - 1) = 0
                oRsDetail.MoveNext
            Loop
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub SaveGrid()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    Dim sAkses As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.login)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(0)

            Dim i As Double
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = -1 Then
                    sAkses = "Y"
                Else
                    sAkses = "N"
                End If
                              
   
                sQuery = "update companyaccess set accessctrl='" & sAkses & "' where emplcode ='" & Text1(0) & "' "
                sQuery = sQuery & " and dbid='" & .TextMatrix(i, 0) & "'"
                                                
                Select Case ToNumber(.TextMatrix(i, .Cols - 1))
                Case 1 And Not .TextMatrix(i, 0) = ""
                        
                        
                       ' oKon.Execute sInsertDetail
                        
                Case 2
                        oKon.Execute sQuery
                        
                Case 3
                        'oKon.Execute sDeleteDetail
                End Select
            Next

        'End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub SaveGrid2()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    Dim sAkses As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.login)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid(1)

            Dim i As Double
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = -1 Then
                    sAkses = "Y"
                Else
                    sAkses = "N"
                End If
                              
   
                sQuery = "update groupaccess set accessctrl='" & sAkses & "' where emplcode ='" & Text1(0) & "' "
                sQuery = sQuery & " and groupid='" & .TextMatrix(i, 0) & "'"
                                                
                Select Case ToNumber(.TextMatrix(i, .Cols - 1))
                Case 1 And Not .TextMatrix(i, 0) = ""
                        
                        
                       ' oKon.Execute sInsertDetail
                        
                Case 2
                        oKon.Execute sQuery
                        
                Case 3
                        'oKon.Execute sDeleteDetail
                End Select
            Next

        'End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

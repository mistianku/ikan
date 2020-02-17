VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form PelajaranLevelGroupFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   8280
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
   ScaleHeight     =   8280
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9615
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   7155
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "PelajaranLevelGroupFrm.frx":0000
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
         Caption         =   "Materi Group"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VSFlex8LCtl.VSFlexGrid oGrid 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   9615
      _cx             =   16960
      _cy             =   11668
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"PelajaranLevelGroupFrm.frx":001C
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
Attribute VB_Name = "PelajaranLevelGroupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim skodelevel As String
Dim snamalevel As String
Dim skodegroup As String
Dim sobjtype As String
Dim saudituser As String
Dim sauditdate As Date
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
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        oSaveDetail
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
        oCon.Close
        DoSaveData = True
        istatus = Normal
        ShowGrid skodegroup
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
        sQuery = "Delete from master_pelajaran_group where kodegroup='" & skodegroup & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
  
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skodegroup = oRubahKutip(Text1(0).Text)
    snamagroup = oRubahKutip(Text1(1).Text)
     
    sUpdate = "update master_pelajaran_group set "
    sUpdate = sUpdate & "namagroup='" & snamagroup & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "kodegroup='" & skodegroup & "'"
    
    sInsert = "insert into master_pelajaran_group ("
    sInsert = sInsert & "kodegroup,namagroup ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & skodegroup & "',"
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
sTitle = "Master Level Materi Belajar"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnLevel

BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
cleardata
istatus = Normal
MoveLast
ShowGrid Text1(0).Text
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


Private Sub Text1_Change(Index As Integer)
ShowGrid Text1(0)
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
Public Sub ShowGrid(skode As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    sKondisi = " Where kodegroup='" & skode & "' Order By kodelevel Asc"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "SELECT kodelevel,namalevel FROM master_pelajaran_level  "
    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid

        .COLWIDTH(1) = .Width - .COLWIDTH(0) - 100
        GridModul.ClearGridDetail oGrid
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("kodelevel"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("namalevel"))
                .TextMatrix(i, .Cols - 1) = 0
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
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

Public Sub AddRow()
With oGrid
If .TextMatrix(.row, 0) = "" Then Exit Sub
    'If .row < .Rows - 1 And .TextMatrix(.row + 1, 0) = "" Then Exit Sub
        If .row < .Rows - 1 Then
           .Select .row + 1, 0
           .EditCell
        Else
            .Rows = .Rows + 1
            .Select .row + 1, 0
            .EditCell
        End If
        .TextMatrix(.row, .Cols - 1) = "1"
End With
End Sub
Private Sub oGrid_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = oGrid.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, oGrid, istatus
With oGrid

    Select Case .col
    Case 0
    
    Case 1
    
    End Select

End With
End Sub
Private Sub ogrid_KeyDown(KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
With oGrid

    'If Not ToNumber(.TextMatrix(.row, .Cols - 1)) = 0 Then Exit Sub
    If Not KeyCode = vbKeyInsert Then
           gridDetail_KeyDown KeyCode, 0, oGrid, istatus
           If KeyCode = vbKeyDelete Then Exit Sub
           Select Case .col
           Case 0
                If ToNumber(.TextMatrix(.row, .Cols - 1)) = 1 Then
                    .EditCell
                    .EditCell
                End If
           Case 1
                .EditCell
                .EditCell
           End Select
          
           'MsgBox "test"

    Else
            .Rows = .Rows + 1
            .Select .Rows - 1, 0
            .EditCell
'           gridDetail_KeyDown KeyCode, 0, oGrid, istatus

    End If
End With
End Sub

Private Sub ogrid_KeyDownEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim sProductIDTmp As String
Dim irow As Integer
With oGrid
Select Case col
    Case 0
                If KeyCode = 13 And ToNumber(.TextMatrix(.row, .Cols - 1)) = 1 Then
                    irow = .row
                    .Select .row, .col
                    sProductIDTmp = Trim(.TextMatrix(.row, .col))
                    .TextMatrix(.row, .col) = ""
                    If .FindRow(sProductIDTmp, , 0, True) = -1 Then
                        .TextMatrix(.row, .col) = sProductIDTmp
                    Else
                        MsgBox "Kode " & .TextMatrix(.FindRow(sProductIDTmp, , 0, True), .col) & " Sudah Ada !! ", vbInformation
                    End If
                End If
                
    Case 1
    
End Select
End With
End Sub
Private Sub oGrid_Click()
With oGrid
    If .Rows = 1 Then
                AddRow
            End If
    Select Case .col
    Case 1
            
    End Select
End With

End Sub
Private Sub oSaveDetail()
Dim irow As Integer
With oGrid
    For irow = 1 To .Rows - 1
        skodelevel = oRubahKutip(.TextMatrix(irow, 0))
        snamalevel = oRubahKutip(.TextMatrix(irow, 1))
        Select Case .TextMatrix(irow, .Cols - 1)
        Case Is = 1 'Insert
                    sQuery = "insert into master_pelajaran_level "
                    sQuery = sQuery & "("
                    sQuery = sQuery & "kodelevel,"
                    sQuery = sQuery & "namalevel,"
                    sQuery = sQuery & "kodegroup,"
                    sQuery = sQuery & "audituser,"
                    sQuery = sQuery & "auditdate"
                    sQuery = sQuery & ") "
                    sQuery = sQuery & "values "
                    sQuery = sQuery & "('"
                    sQuery = sQuery & skodelevel & "','"
                    sQuery = sQuery & snamalevel & "','"
                    sQuery = sQuery & skodegroup & "','"
                    sQuery = sQuery & MenuFrm.sUserID & "','"
                    sQuery = sQuery & Format(Now(), "YYYY-MM-DD") & "'"
                    sQuery = sQuery & ")"
                    oCon.Execute sQuery
        Case Is = 2 'Update
                    sQuery = "Update master_pelajaran_level "
                    sQuery = sQuery & "set namalevel='" & snamalevel & "'"
                    sQuery = sQuery & " where kodelevel='" & skodelevel & "' and kodegroup='" & skodegroup & "'"
                    oCon.Execute sQuery
        Case Is = 3 'Delete
                    sQuery = "Delete From master_pelajaran_level "
                    sQuery = sQuery & " where kodelevel='" & skodelevel & "' and kodegroup='" & skodegroup & "'"
                    oCon.Execute sQuery
        End Select
    Next
End With
End Sub


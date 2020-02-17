VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form GLKwitansiToGL 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Master Data User"
   ClientHeight    =   6855
   ClientLeft      =   22080
   ClientTop       =   3450
   ClientWidth     =   13260
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
   ScaleHeight     =   6855
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Top             =   2760
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tidak Pilih Semua"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   8
         Top             =   2760
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pilih Semua"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   2760
         Width           =   2295
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
         Left            =   2700
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3120
         Visible         =   0   'False
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
         Left            =   2700
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLKwitansiToGL.frx":0000
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
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   11535
         _cx             =   20346
         _cy             =   2778
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12632256
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
         Rows            =   4
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"GLKwitansiToGL.frx":001C
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Agama"
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Kode"
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GLKwitansiToGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim iModulku As Modul

Dim skode As String
Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from master_agama where kode='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
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
    sQuery = "Select *  from master_agama order by kode asc limit 1"
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
    sQuery = "Select  *  from master_agama where kode >'" & Text1(0).Text & "' order by kode asc limit 1"
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
    sQuery = "Select  *  from master_agama where kode<'" & Text1(0).Text & "' order by kode desc limit 1"
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
    sQuery = "Select *  from master_agama order by kode desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
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
        sQuery = "Delete from master_agama where kode='" & skode & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skode = ToText(Text1(0).Text)
    sagama = ToText(Text1(1).Text)
     
    sUpdate = "update master_agama set "
    sUpdate = sUpdate & "agama='" & sagama & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "kode='" & skode & "'"
    
    sInsert = "insert into master_agama ("
    sInsert = sInsert & "kode,agama ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & skode & "',"
    sInsert = sInsert & "'" & sagama & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsAgama, ""
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub Command1_Click()
oinsert_transaksi_kwitansi_temp
MsgBox "Proses Transfer ke GL , Complete ", vbInformation
ShowGrid1
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Transfer Kwitansi ke GL "
''lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransferKwitansikeGL
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
MenuFrm.Picture3.Visible = False
End Sub

Private Sub Form_Load()
oFormatOption 1, Me
cleardata
iModulku = mnTransferKwitansikeGL
istatus = Normal
ShowGrid1
'MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).Text = oRs("kode")
    KodeUserAksesTemp = oRs("kode")
    Text1(0).Locked = True
    Text1(1).Text = oRs("agama")
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

Private Sub oGrid1_Click()
With oGrid1
Select Case .col
Case 0
    .EditCell
End Select
End With
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    oPilihCell "Y"
Case 1
    oPilihCell "N"
End Select
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

Public Sub oPilihCell(sPilih As String)
With oGrid1
Dim irow As Integer
For irow = 1 To .Rows - 1
    If sPilih = "Y" Then
    .TextMatrix(irow, 0) = -1
    Else
    .TextMatrix(irow, 0) = 0
    End If
Next
End With
End Sub
Public Sub ShowGrid1()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    sQuery = "call spget_transaksi_kwitansi_to_gl_view"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        .Cols = 9
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("yop"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("mop"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("txtjnsbyr"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("jmldok"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("totalstlpfaktur"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("txtdokstatus"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("jnsbayar"))
                .TextMatrix(i, .Cols - 1) = RTrim(oRsDetail("dokstatus"))
                '.TextMatrix(i, .Cols - 1) = RTrim(oRsDetail("docentry"))
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
           
            '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub oinsert_transaksi_kwitansi_temp()
On Error GoTo errhandler
Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim saudituser  As String
    Dim syop  As Integer
    Dim smop  As Integer
    Dim sjnsbayar  As String
    saudituser = MenuFrm.sUserID
    sQuery = "delete from transaksi_kwitansi_temp where audituser='" & saudituser & "'"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    oKon.Execute (sQuery)
    With oGrid1
    Dim irow As Integer
    For irow = 1 To .Rows - 1
            If .TextMatrix(irow, 0) = -1 And .TextMatrix(irow, .Cols - 1) = "O" Then
                
                syop = .TextMatrix(irow, 1)
                smop = .TextMatrix(irow, 2)
                sjnsbayar = .TextMatrix(irow, .Cols - 2)
                sQuery = "call spinsert_transaksi_kwitansi_temp('"
                sQuery = sQuery & saudituser & "','"
                sQuery = sQuery & syop & "','"
                sQuery = sQuery & smop & "','"
                sQuery = sQuery & sjnsbayar & "','"
                sQuery = sQuery & GetDocnum(transaksi_trngl, True, DBaseConection.Modul) & "')"
                oKon.Execute (sQuery)
                
            End If
        
    Next
    End With
    sQuery = "call sp_proses_kwitansi_to_gl('" & Format(Now(), "YYYY-MM-DD") & "','" & "" & "','" & MenuFrm.sUserID & "')"
    oKon.Execute (sQuery)
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub Execution()

End Sub

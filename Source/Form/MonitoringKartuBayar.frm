VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form MonitoringKartuBayar 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   8220
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   12375
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
   ScaleHeight     =   8220
   ScaleWidth      =   12375
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr1 
      Left            =   10800
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSDFLATS.FlatButton BrowseUserID 
      Height          =   255
      Left            =   9000
      TabIndex        =   13
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      MouseIcon       =   "MonitoringKartuBayar.frx":0000
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Kartu Bayar"
      Height          =   3615
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid oGrid2 
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   5741
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
         Rows            =   50
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringKartuBayar.frx":001C
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "K e l a s"
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   3201
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
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringKartuBayar.frx":00CA
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Periode"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   11775
      Begin VSDFLATS.FlatComboBox fcTahun 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         MouseIcon       =   "MonitoringKartuBayar.frx":0137
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
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
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
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
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   6675
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
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   7155
   End
   Begin VB.Label Label1 
      Caption         =   "A l a m a t"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "No.ID.Siswa"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "MonitoringKartuBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim snoidsiswa As String
Dim snmlengkap As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from master_siswa where noidsiswa='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringKartuSiswa
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
    sQuery = "Select *  from master_siswa order by noidsiswa asc limit 1"
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
    sQuery = "Select  *  from master_siswa where noidsiswa >'" & Text1(0).Text & "' order by noidsiswa asc limit 1"
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
    sQuery = "Select  *  from master_siswa where noidsiswa<'" & Text1(0).Text & "' order by noidsiswa desc limit 1"
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
    sQuery = "Select *  from master_siswa order by noidsiswa desc limit 1 "
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
'Dim ires As Integer
'    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
'    If ires = 6 Then
'        If DoSaveData Then
'             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
'             MenuFrm.SetToolbarku me , istatus, MenuFrm.sGroupUserID, mnMonitoringPembayaran
'        End If
'    End If
End Sub
Public Sub DeleteData()
'    Dim ires As Integer
'    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
'    If ires = 6 Then
'        If DoDeleteData Then
'             MsgBox "Data Sudah Terhapus", , "Hapus Data"
'             MovePrevious
'        End If
'    End If
'    MenuFrm.SetToolbarku me , istatus, MenuFrm.sGroupUserID, mnMonitoringPembayaran
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
        sQuery = "Delete from master_siswa where noidsiswa='" & snoidsiswa & "'"
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
'    KodeUserAksesTemp = Text1(0)
'    istatus = StatusForm.DataBaru
'    cleardata
'    MenuFrm.SetToolbarku me , istatus, MenuFrm.sGroupUserID, mnMonitoringPembayaran
'    Text1(0).Locked = False
'    Text1(0).SetFocus
'    Text1(0).TabIndex = 0
'    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringKartuSiswa
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    snoidsiswa = ToText(Text1(0).Text)
    snmlengkap = ToText(Text1(1).Text)
     
    sUpdate = "update master_siswa set "
    sUpdate = sUpdate & "nmlengkap='" & snmlengkap & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "'"
    
    sInsert = "insert into master_siswa ("
    sInsert = sInsert & "noidsiswa,nmlengkap ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & snoidsiswa & "',"
    sInsert = sInsert & "'" & snmlengkap & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsSiswa, ""
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub fcTahun_Click()
ShowGrid1 Text1(0)
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = " Monitoring Kartu Pembayaran Siswa "
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringKartuSiswa

BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
cleardata
oAddfcTahun
istatus = Normal
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    GridModul.ClearGridDetail ogrid1
    GridModul.ClearGridDetail ogrid2
    Text1(0).Text = ToText(oRs("noidsiswa"))
    KodeUserAksesTemp = oRs("noidsiswa")
    Text1(0).Locked = True
    Text1(1).Text = oRs("nmlengkap")
    Text1(2).Text = oRs("almtrumah1")
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
With ogrid1
If .row = 0 Then Exit Sub
ShowGrid2 .TextMatrix(.row, .Cols - 1), ToNumber(fcTahun.Text)
End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 0
    ShowGrid1 ToText(Text1(0).Text)
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
If Index = 0 Then FindData ToText(Text1(0).Text)
End Sub

Public Sub oAddfcTahun()
On Error GoTo errhandler
Dim i As Integer
        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "SELECT  MIN(yop) AS yop1,MAX(yop) AS yop2 FROM master_kartu_bayar"
        Set oRs = oCon.Execute(sQuery)
        For i = oRs(0) To oRs(1)
            fcTahun.AddItem i
        Next
            fcTahun.Text = Year(Now)
        oCon.Close
        Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Add Combo Box Tahun"
End Sub




Public Sub ShowGrid1(keynoidsiswa As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    sQuery = "    SELECT a.docentry,a.nokursus,a.tglmulai,ifnull(b.keterangan,'-') AS kelas ,a.noidsiswa  FROM master_kelas AS a "
    sQuery = sQuery & "   LEFT JOIN master_default_pelajaran AS b ON a.pelajaran=b.pelajaran "
    sQuery = sQuery & "   WHERE a.noidsiswa='" & keynoidsiswa & "' and stskelas='1'"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid1

        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid1
        .Cols = 4
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("nokursus"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("tglmulai"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("kelas"))
                .TextMatrix(i, .Cols - 1) = RTrim(oRsDetail("docentry"))
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
            ShowGrid2 .TextMatrix(.row, .Cols - 1), ToNumber(fcTahun.Text)
            '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub ShowGrid2(keyDocentry As String, keyYop As Integer)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    sQuery = "    SELECT * FROM vmaster_kartu_bayar "
    sQuery = sQuery & "   WHERE docentry = " & keyDocentry & " AND yop = " & keyYop
    sQuery = sQuery & "   Order by mop asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid2
        .Cols = 6
        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid2
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                
                .TextMatrix(i, 0) = ToText(oRsDetail("namabulan"))
                .TextMatrix(i, 1) = ToText(oRsDetail("tglbayar"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("nokwitansi"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("bayar"))
                .TextMatrix(i, 4) = formatRupiah(oRsDetail("nilaibayar"))
                .TextMatrix(i, .Cols - 1) = ToText(oRsDetail("stsbayar"))
                
'                If .TextMatrix(i, .Cols - 1) = "0" Then
'                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = &H8000000A       '&H8000000A&
'                Else
'                    .Cell(flexcpFontBold, i, 0, , .Cols - 1) = True
'                    .Cell(flexcpForeColor, i, 0, , .Cols - 1) = vbRed
'                    .Cell(flexcpBackColor, i, 0, , .Cols - 1) = vbGreen
'                End If
                oRsDetail.MoveNext
            Loop
            If .Rows - 1 = 12 Then
               .Select Month(Date), 0
               '.Cell(flexcpBackColor, Month(Date), 0, , .Cols - 1) = vbGreen
            End If
            
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim snokursusQ As String
Me.CR1.Reset
Me.CR1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.CR1.ReportFileName = App.Path + "\Reports\Kartu_Bayar_Rpt.Rpt"

Dim sKriteria As String

sKriteria = " where nokursus  between '" & ogrid1.TextMatrix(ogrid1.row, 0) & "' and '" & ogrid1.TextMatrix(ogrid1.row, 0) & "'"

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vkartu_pembayaran_rpt vkartu_pembayaran_rpt1" & sKriteria
'
'
Me.CR1.SQLQuery = sQuery
Me.CR1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.CR1.ParameterFields(1) = "cmpnyname" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.CR1.ParameterFields(2) = "address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.CR1.ParameterFields(3) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
'Me.CR1.ParameterFields(1) = "@Priceid2" & ";" & Text1(17).Text & ";" & True

Me.CR1.Destination = crptToWindow
Me.CR1.RetrieveDataFiles
Me.CR1.WindowState = crptMaximized
Me.CR1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form SettingLockSalesFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Setting Lock Sales  Form"
   ClientHeight    =   8355
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   12255
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
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   6615
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   12855
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   4695
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   12615
         _cx             =   22251
         _cy             =   8281
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
         BackColorSel    =   16767152
         ForeColorSel    =   198
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16382457
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"SettingLockSales.frx":0000
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
      Caption         =   "Cari Customer berdasarkan"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   12855
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
         Left            =   3120
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   360
         Width           =   8595
      End
      Begin VB.Label Label1 
         Caption         =   "Kode atau Nama Customer"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "Lock Sales"
      ForeColor       =   &H80000008&
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   12855
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "unLock"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Lock"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   9240
         TabIndex        =   11
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         FlatButton      =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Mulai Lock"
         Height          =   315
         Index           =   3
         Left            =   7080
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   11895
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
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SettingLockSales.frx":00F5
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
         Caption         =   "Urutan Tampil"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Area"
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Area"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "SettingLockSalesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim sText As String
Dim slock_sales As String
Dim slock_start_date As String
Dim skodecustomer As String
Dim skodearea As String
Dim snamaarea As String
Dim svisorder As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub Execution()

End Sub
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from master_area where kodearea='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
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
    sQuery = "Select *  from master_area order by kodearea asc limit 1"
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
    sQuery = "Select  *  from master_area where kodearea >'" & Text1(0).text & "' order by kodearea asc limit 1"
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
    sQuery = "Select  *  from master_area where kodearea<'" & Text1(0).text & "' order by kodearea desc limit 1"
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
    sQuery = "Select *  from master_area order by kodearea desc limit 1 "
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
             FindData skodearea
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
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
        sQuery = "Delete from master_area where kodearea='" & skodearea & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1

End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    skodearea = ToText(Text1(0).text)
    snamaarea = ToText(Text1(1).text)
    svisorder = ToText(Text1(2).text)
     
    svisorder = IIf(svisorder = "", skodearea, svisorder)
    sUpdate = "update master_area set "
    sUpdate = sUpdate & "namaarea='" & snamaarea & "',visorder='" & svisorder & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "kodearea='" & skodearea & "'"
    
    sInsert = "insert into master_area ("
    sInsert = sInsert & "kodearea,namaarea,visorder ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & skodearea & "',"
    sInsert = sInsert & "'" & snamaarea & "',"
    sInsert = sInsert & "'" & svisorder & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsArea, "", ubAscending, DBaseConection.Modul
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Setting Lock Sales Form"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSettingLockSalesFrm
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, typemenu.entrian, MenuFrm.sinsertmodul
oFormatOption 1, Me
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
cleardata
istatus = Normal
FlatDatePicker1.value = Now()
ShowGrid
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("kodearea")
    KodeUserAksesTemp = oRs("kodearea")
    Text1(0).Locked = True
    Text1(1).text = oRs("namaarea")
    Text1(2).text = oRs("visorder")
    'Text1(2).Text = DecryptPassword(oRs("Password"))
    'Me.Caption = DecryptPassword(oRs("Password"))
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    Text1(2).Enabled = True
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

Private Sub oGrid1_Click(Index As Integer)
With oGrid1(0)
    Select Case .col
    Case 0
            .EditCell
            If .TextMatrix(.row, 0) = "-1" Then
                .TextMatrix(.row, 1) = Format(FlatDatePicker1.value, "YYYY-MM-DD")
                slock_sales = 1
                
            Else
                .TextMatrix(.row, 1) = Format("", "YYYY-MM-DD")
                slock_sales = 0
            End If
            skodecustomer = .TextMatrix(.row, 2)
            DoSaveDataLogSales
            
   End Select
End With
End Sub

Private Sub Option1_Click(Index As Integer)
ShowGrid
End Sub

Private Sub Text1_Change(Index As Integer)
ShowGrid
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
Public Sub ShowGrid()
On Error GoTo errhandler
Dim stukarfaktur As String
Dim sskeyfind As Integer
Dim skata As String
Dim skeytgl As Integer
Dim stgl As Date

sText = "%" & Text1(3) & "%"
slock_sales = IIf(Option1(0).value = True, 1, 0)
slock_start_date = FlatDatePicker1.value
'stukarfaktur CHAR(1),skeyfind INT,skata VARCHAR(100),
'skeytgl INT ,stgl DATE

    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "CALL sp_lock_sales_view('"
    sQuery = sQuery & sText & "','"
    sQuery = sQuery & slock_sales & "')"
   

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(0)

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1(0)
        .ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("lock_sales"))
                
                .TextMatrix(i, 1) = Format(IIf(RTrim(oRsDetail("lock_sales")) = 0, "", RTrim(oRsDetail("lock_start_date"))), "YYYY-MM-DD")
'                IIf(RTrim(oRsDetail("lock_start_date")) = Null, "2018-01-01", RTrim(oRsDetail("lock_start_date")))
                .TextMatrix(i, 2) = RTrim(oRsDetail("kodecustomer"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("namacustomer"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("alamat"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("kota"))
                .TextMatrix(i, .Cols - 1) = 0
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Private Function DoSaveDataLogSales() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        
        sQuery = "update master_customer set lock_sales='" & slock_sales & "',lock_start_date='" & Format(slock_start_date, "YYYY-MM-DD") & "'"
        sQuery = sQuery & " where kodecustomer='" & skodecustomer & "'"
        oCon.Execute sQuery
        oCon.Close
        DoSaveDataLogSales = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "savedata"
End Function

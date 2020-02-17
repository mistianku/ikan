VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MonitoringTukarFakturBrowse 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   8055
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   14580
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
   ScaleHeight     =   8055
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Tidak Pilih Semua"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   22
      Top             =   7440
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Pilih Semua"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1095
      Index           =   3
      Left            =   7080
      TabIndex        =   16
      Top             =   840
      Width           =   7455
      Begin VSDFLATS.FlatButton FlatButton1 
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         MouseIcon       =   "MonitoringTukarFakturBrowse.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Update"
      End
      Begin NeoCalendarII.DatePicker DatePicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
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
         CalendarPresentDateColor=   -2147483646
      End
      Begin VSDFLATS.FlatButton FlatButton1 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         MouseIcon       =   "MonitoringTukarFakturBrowse.frx":001C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ambil"
      End
      Begin VSDFLATS.FlatButton FlatButton1 
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   20
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         MouseIcon       =   "MonitoringTukarFakturBrowse.frx":0038
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Batal"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Dokumen Tukar Faktur"
      Height          =   5295
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   14295
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   4935
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   14055
         _cx             =   24791
         _cy             =   8705
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringTukarFakturBrowse.frx":0054
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
      Caption         =   "Pencarian Dokumen"
      Height          =   1095
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   6735
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   600
         Width           =   4335
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
         MouseIcon       =   "MonitoringTukarFakturBrowse.frx":01EA
         Enabled         =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Kata Kunci"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cari Berdasarkan"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Berdasarkan Tanggal Dokumen"
      Height          =   615
      Index           =   0
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin NeoCalendarII.DatePicker DatePicker1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
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
         CalendarPresentDateColor=   -2147483646
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Semua"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sesudah"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sebelum"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Tukar Faktur"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Semua Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Sudah"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Belum"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "MonitoringTukarFakturBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim skodebrand As String
Dim snamabrand As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Dim sawal  As Integer
Dim oFormAsal As Form
Dim oGridasal As VSFlexGrid
Dim slinenumasal As Integer
Dim sdocentryasal As Integer


Public Sub SaveData()
Dim ires As Integer
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnBrandFrm
        End If
    End If
End Sub

Private Function DoSaveData() As Boolean
On Error GoTo errhandler
SimpanGrid1
ShowGrid
'    If setData Then
'        If oCon.State = 1 Then oCon.Close
'         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
'        If istatus = StatusForm.DataBaru Then
'        sQuery = sInsert
'        Else
'        sQuery = sUpdate
'        End If
'        oCon.Execute sQuery
'        oCon.Close
'        DoSaveData = True
'        istatus = Normal
'        Exit Function
'    End If
DoSaveData = True
Exit Function
errhandler:
MainModule.ShowMessage Err.Description, "savedata"
End Function
Private Function DoDeleteData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "Delete from master_brand where kodebrand='" & skodebrand & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnBrandFrm
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub


Private Function setData() As Boolean
On Error GoTo errhandler
    skodebrand = ToText(Text1(0).text)
    snamabrand = ToText(Text1(1).text)
     
    sUpdate = "update master_brand set "
    sUpdate = sUpdate & "namabrand='" & snamabrand & "' where "
    'sUpdate = sUpdate & " where "
    sUpdate = sUpdate & "kodebrand='" & skodebrand & "'"
    
    sInsert = "insert into master_brand ("
    sInsert = sInsert & "kodebrand,namabrand ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & skodebrand & "',"
    sInsert = sInsert & "'" & snamabrand & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub DatePicker1_Change(Index As Integer)
If Index = 0 Then
    If sawal = 1 Then
        ShowGrid
    End If
End If
End Sub

Private Sub FlatButton1_Click(Index As Integer)
Select Case Index
Case 1
    oTaruhData
    Unload Me
Case 2
    Unload Me
End Select
End Sub

'Private Sub FlatButton1_Click()
''oGrid1(0).TextMatrix(oGrid1(0).row, 2) = DatePicker1(1).value
'End Sub

Private Sub FlatComboBox1_Click()
If sawal = 1 Then
ShowGrid
End If
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Monitoring Entry Tukar Faktur"

Me.Caption = " " & sTitle & " "
'MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnBrand

End Sub

Private Sub Form_Load()
sawal = 0
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 3, Me
cleardata
istatus = Normal
FlatComboBox1.AddItem "Nama Customer"
FlatComboBox1.AddItem "Kode Customer"
FlatComboBox1.AddItem "No. Dokumen"
FlatComboBox1.ListIndex = 0

DatePicker1(0).value = Now()
DatePicker1(1).value = Now()
Text1(0).Locked = False

'ShowGrid
sawal = 1
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
        If .TextMatrix(.row, 1) = 1 Then
            .EditCell
        End If
        
    Case 1
'        If .TextMatrix(.row, 0) = -1 Then
'            .EditCell
'        End If
'        If .TextMatrix(.row, 1) = -1 Then
'            .TextMatrix(.row, 2) = DatePicker1(1).value
'        Else
'            .TextMatrix(.row, 2) = ""
'        End If
    End Select
End With
End Sub

Private Sub Option1_Click(Index As Integer)
If sawal = 1 Then
ShowGrid
End If
End Sub

Private Sub Option2_Click(Index As Integer)
ShowGrid
End Sub

Private Sub Option3_Click(Index As Integer)
Select Case Index
Case 0
    oPilihGrid True
Case 1
    oPilihGrid False
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 0 Then
    If sawal = 1 Then
        ShowGrid
    End If
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub


Public Sub ShowGrid()
On Error GoTo errhandler
Dim stukarfaktur As String
Dim sskeyfind As Integer
Dim skata As String
Dim skeytgl As Integer
Dim stgl As Date

stukarfaktur = IIf(Option1(0).value, "N", IIf(Option1(1).value, "Y", "A"))
sskeyfind = FlatComboBox1.ListIndex + 1
skata = "%" & Text1(0) & "%"
skeytgl = IIf(Option2(0).value, 1, IIf(Option2(1).value, 2, 3))
stgl = DatePicker1(0).value
'stukarfaktur CHAR(1),skeyfind INT,skata VARCHAR(100),
'skeytgl INT ,stgl DATE

    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
      
    sQuery = "CALL sp_transaksi_keluar_tfaktur_get_kwitansi('"
    sQuery = sQuery & stukarfaktur & "','"
    sQuery = sQuery & sskeyfind & "','"
    sQuery = sQuery & skata & "','"
    sQuery = sQuery & skeytgl & "','"
    sQuery = sQuery & Format(stgl, "YYYY-MM-DD") & "')"
   

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
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = IIf(oRsDetail("tukarfaktur") = "N", 0, 1)
                .TextMatrix(i, 2) = IIf(IsNull(oRsDetail("tgltukarfaktur")), "", oRsDetail("tgltukarfaktur"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("nodokumen"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("tgldokumen"))
                .TextMatrix(i, 5) = IIf(IsNull(oRsDetail("tgljtfaktur")), "", oRsDetail("tgljtfaktur"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("kodecustomer"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("namacustomer"))
                .TextMatrix(i, 8) = RTrim(oRsDetail("totalsetppn"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("docentry"))
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
Public Sub SimpanGrid1()
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim irow As Integer
    Dim snodokumen As String
    Dim stukarfaktur As String
    Dim stgltukarfaktur As String
    
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    With oGrid1(0)
        For irow = 1 To .Rows - 1
            stukarfaktur = IIf(.TextMatrix(irow, 1) = -1, "Y", "N")
            stgltukarfaktur = ToText(.TextMatrix(irow, 2))
            snodokumen = .TextMatrix(irow, 3)
            sQuery = "call sp_transaksi_keluar_tfaktur_update('" & snodokumen & "','"
            sQuery = sQuery & stukarfaktur & "','"
            sQuery = sQuery & Format(stgltukarfaktur, "YYYY-MM-DD") & "')"
            If .TextMatrix(irow, 0) = -1 Then
                oKon.Execute (sQuery)
            End If
        Next
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Save Produk Diskon"
End Sub
Public Sub ShowForm(ogrid As VSFlexGrid, sdocentry As Double, slinenum As Integer, skodecustomer As String)
    Set oGridasal = ogrid
    sdocentryasal = sdocentry
    slinenumasal = slinenum
    Text1(0) = skodecustomer
    FlatComboBox1.ListIndex = 1
    Me.Show 1
    
End Sub

Public Sub oPilihGrid(sPilih As Boolean)
With oGrid1(0)
    Dim i As Integer
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 1) = 1 Then
            If sPilih Then
                .TextMatrix(i, 0) = -1
            Else
                .TextMatrix(i, 0) = 0
            End If
        End If
    Next
    .Refresh
End With
End Sub



Public Sub oTaruhData()
Dim slineQ As Integer
Dim i As Integer

With oGrid1(0)
    For i = 1 To .Rows - 1
    If .TextMatrix(i, 0) = -1 Then
        
        oGridasal.Rows = oGridasal.Rows + 1
        oGridasal.TextMatrix(oGridasal.Rows - 1, 0) = .TextMatrix(i, 2)
        oGridasal.TextMatrix(oGridasal.Rows - 1, 1) = .TextMatrix(i, 3)
        oGridasal.TextMatrix(oGridasal.Rows - 1, 2) = .TextMatrix(i, 4)
        oGridasal.TextMatrix(oGridasal.Rows - 1, 3) = .TextMatrix(i, 5)
        oGridasal.TextMatrix(oGridasal.Rows - 1, 4) = .TextMatrix(i, 8)
        oGridasal.TextMatrix(oGridasal.Rows - 1, oGridasal.Cols - 4) = sdocentryasal
        oGridasal.TextMatrix(oGridasal.Rows - 1, oGridasal.Cols - 3) = slinenumasal
        oGridasal.TextMatrix(oGridasal.Rows - 1, oGridasal.Cols - 2) = .TextMatrix(i, .Cols - 2)
        oGridasal.TextMatrix(oGridasal.Rows - 1, oGridasal.Cols - 1) = 1
        oGridasal.Cell(flexcpFontBold, oGridasal.Rows - 1, 0, , .Cols - 1) = vbNormal
        slinenumasal = slinenumasal + 1
    End If
    Next
    .Refresh
    KwitansiFrm.lbllinenum = slinenumasal
    KwitansiFrm.oRecalculate
End With
End Sub

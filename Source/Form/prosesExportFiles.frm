VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form prosesExportFiles 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Backup Database Form"
   ClientHeight    =   9345
   ClientLeft      =   22080
   ClientTop       =   3450
   ClientWidth     =   12330
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
   ScaleHeight     =   9345
   ScaleWidth      =   12330
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Info Dokumen"
      Height          =   1215
      Index           =   6
      Left            =   4680
      TabIndex        =   25
      Top             =   120
      Width           =   7335
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Total Nilai"
         Height          =   315
         Index           =   4
         Left            =   3120
         TabIndex        =   27
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah "
         Height          =   315
         Index           =   3
         Left            =   3120
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   5
      Left            =   120
      TabIndex        =   20
      Top             =   6960
      Width           =   11775
      Begin VB.OptionButton Option2 
         Caption         =   "Tidak Pilih Semua"
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pilih Semua"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Penjualan"
      Height          =   1095
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   4455
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Enabled         =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
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
         Enabled         =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Dari"
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai"
         Height          =   315
         Index           =   2
         Left            =   840
         TabIndex        =   16
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Export Status"
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton Option1 
         Caption         =   "All"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Close"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Open"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Backup Otomatis"
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   8640
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Backup Pada Saat Tutup Aplikasi"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Backup Pada Saat Aktif Aplikasi"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4695
      End
   End
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   7680
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "prosesExportFiles.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Proses Export Files"
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   12240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "txt"
      InitDir         =   "c:\*.SQL"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
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
         Index           =   0
         Left            =   2340
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   8775
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   11160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "prosesExportFiles.frx":001C
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
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ke Direktori "
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   8160
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "prosesExportFiles.frx":0038
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Copy Export Files"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Daftar File Export"
      Height          =   5055
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   11895
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   4575
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   11655
         _cx             =   20558
         _cy             =   8070
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
         BackColorSel    =   8454143
         ForeColorSel    =   4194432
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"prosesExportFiles.frx":0054
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
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   24
      Top             =   1440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      MouseIcon       =   "prosesExportFiles.frx":0121
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tampilkan Data"
   End
End
Attribute VB_Name = "prosesExportFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim iModulku As Modul
Dim sid As Integer
Dim sfile_backup As String
Dim sbackup_login_aplikasi As String
Dim sbackup_exit_aplikasi As String
Dim sawal As String
Dim saudituser As String
Dim sauditdate As String

Dim sexpsts As String
Dim salltgl As String
Dim stglfr As String
Dim stglto As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from setting_backup_database where id='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
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
    sQuery = "Select *  from setting_backup_database order by id asc limit 1"
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
    sQuery = "Select  *  from setting_backup_database where id >'" & Text1(0).text & "' order by id asc limit 1"
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
    sQuery = "Select  *  from setting_backup_database where id<'" & Text1(0).text & "' order by id desc limit 1"
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
    sQuery = "Select *  from setting_backup_database order by id desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
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
        sQuery = "Delete from setting_backup_database where id='" & sid & "'"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    sid = 1
    
    Text1(0) = IIf(InStr(1, Text1(0), ".") = 0, cd1.FileName, Left(Text1(0), InStr(1, Text1(0), ".")) & "SQL")
    
    sfile_backup = Replace(Text1(0).text, "\", "\\")
    If Check1(0).value = 1 Then
        sbackup_login_aplikasi = "Y"
    Else
        sbackup_login_aplikasi = "N"
    End If
    If Check1(1).value = 1 Then
        sbackup_exit_aplikasi = "Y"
    Else
        sbackup_exit_aplikasi = "N"
    End If
    sUpdate = " update setting_backup_database set "
    sUpdate = sUpdate & " file_backup= '" & sfile_backup & "',"
    sUpdate = sUpdate & " backup_login_aplikasi= '" & sbackup_login_aplikasi & "',"
    sUpdate = sUpdate & " backup_exit_aplikasi= '" & sbackup_exit_aplikasi & "',"
    sUpdate = sUpdate & " audituser= '" & MenuFrm.sUserID & "',"
    sUpdate = sUpdate & " auditdate= '" & Format(Now(), "YYYY-MM-DD") & "'"
    sUpdate = sUpdate & " where id=1"

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
cd1.ShowSave
If cd1.FileName = "*.sql" Then Exit Sub
Text1(0) = IIf(InStr(1, cd1.FileName, ".") = 0, cd1.FileName, Left(cd1.FileName, InStr(1, cd1.FileName, ".")))
Text1(0) = Text1(0) & "SQL"
cd1.FileName = "*.sql"
'MsgBox "Proses Import Data Preference Selesai ", vbInformation
End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
    FlatDatePicker1(0).Enabled = False
    FlatDatePicker1(1).Enabled = False
Else
    FlatDatePicker1(0).Enabled = True
    FlatDatePicker1(1).Enabled = True
End If
End Sub

Private Sub FlatButton1_Click(Index As Integer)
Dim sNamaFile As String
Dim a As Integer
Dim sfileheader As String
Dim sfiledetail As String
Dim sthn, sbln, stgl As String
Dim stglku As String
' stglku=right(year(now()),2) & right(100+month(now()),2) & right
stglku = Format(Now(), "YYMMDD") & "_sales.txt"

sfileheader = App.Path & "\exportfilestemp\" & "sales.txt"
sfiledetail = App.Path & "\exportfilestemp\" & "sales_detail1.txt"

Select Case Index
Case 0
    oSimpanGrid
    oExportFile
      
    FlatButton1(1).Enabled = True

Case 1

cd1.Filter = "txt"
cd1.FileName = stglku
cd1.ShowSave
        If cd1.FileName = "*.txt" Then Exit Sub
    sNamaFile = Replace(cd1.FileName, ".txt", "") & ".txt"
    CopyFileBackup sfileheader, sNamaFile
    CopyFileBackup sfiledetail, Replace(sNamaFile, ".txt", "_detail1.txt")
    MsgBox "Copy File Export Complete", vbInformation, "Proses Copy File Export"
    FlatButton1(1).Enabled = False
Case 2
    
    
    stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
    stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")
    salltgl = IIf(Check2 = 1, "A", "N")
    sexpsts = IIf(Option1(0) = True, "1", IIf(Option1(1) = True, "2", "3"))
    ShowGrid sexpsts, salltgl, stglfr, stglto

End Select
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Proses Export Files"

Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesExportFiles
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
MenuFrm.Picture3.Visible = False
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 2, Me
FlatButton1(1).Enabled = False
cleardata
MoveFirst
FlatDatePicker1(0).value = DateSerial(Year(Now()), Month(Now()), 1)
FlatDatePicker1(1).value = Now()
stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")
salltgl = IIf(Check2 = 1, "A", "N")
sexpsts = IIf(Option2(0) = True, "1", IIf(Option2(1) = True, "2", "3"))

ShowGrid sexpsts, salltgl, stglfr, stglto
iModulku = mnProsesExportFiles
istatus = Normal

'Text1(0) = App.Path & "BackupDatabase"
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("file_backup")
    KodeUserAksesTemp = oRs("id")
    'Text1(0).Locked = True

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    If oRs("backup_login_aplikasi") = "Y" Then
        Check1(0).value = 1
    Else
         Check1(0).value = 0
    End If
    If oRs("backup_exit_aplikasi") = "Y" Then
        Check1(1).value = 1
    Else
         Check1(1).value = 0
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

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
    oPilihGrid True
Case 1
    oPilihGrid False
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
If Index = 0 Then FindData Text1(0).text
End Sub




Public Function CopyFileBackup(sFileBackupAsal As String, sFileBackupTujuan As String)
On Error GoTo errhandler
FileCopy sFileBackupAsal, sFileBackupTujuan
' MsgBox "Copy File Export Selesai ", vbInformation
CopyFileBackup = True
Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, " Copy File Backup "
End Function

Public Sub ShowGrid(sexport_sts As String, salltgldokumen As String, stgl_fr As String, stgl_to As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim sjmlfaktur As Double
    Dim snilfaktur As Double
      
    sQuery = " CALL sp_transaksi_keluar_export("
    sQuery = sQuery & "'" & sexport_sts & "','"
    sQuery = sQuery & salltgldokumen & "','"
    sQuery = sQuery & stgl_fr & "','"
    sQuery = sQuery & stgl_to & "')"
'    sQuery = sQuery & "   WHERE kodeproduk='" & keyKodeProduk & "' order by kodediskon asc"

    sjmlfaktur = 0
    snilfaktur = 0
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'    sQuery = "SELECT kodelevelno,namalevelno,nolvlmulai,nolvlselesai,1 as aktif,0 as status FROM master_pelajaran_level_detail  "
'    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

'        .COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid
        .Cols = .Cols + 2
        .ColHidden(.Cols - 1) = True
        .ColHidden(.Cols - 2) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, 0) = RTrim(oRsDetail("pilih"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("nodokumen"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("tgldokumen"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("namacustomer"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("totalsetppn"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("expsts"))
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("docentry"))
                .TextMatrix(i, .Cols - 1) = 0
                sjmlfaktur = sjmlfaktur + 1
                snilfaktur = snilfaktur + ToNumber(.TextMatrix(i, 4))
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
            Text1(1) = formatRupiah(sjmlfaktur)
            Text1(2) = formatRupiah(snilfaktur)
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub oSimpanGrid()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
            sQuery = "delete from transaksi_keluar_export where audituser='" & MenuFrm.sUserID & "'"
            oCon.Execute sQuery
        
        
    
    
    
With ogrid
    Dim i As Integer
    For i = 1 To .Rows - 1
       
            If .TextMatrix(i, 0) = -1 Then
                sQuery = "call sp_transaksi_keluar_export_insert('"
                sQuery = sQuery & .TextMatrix(i, .Cols - 2) & "','"
                sQuery = sQuery & .TextMatrix(i, 1) & "','"
                sQuery = sQuery & MenuFrm.sUserID & "')"
                oCon.Execute sQuery
            End If

    Next
    
End With
oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub
Public Sub oPilihGrid(sPilih As Boolean)
With ogrid
    Dim i As Integer
    For i = 1 To .Rows - 1
       
            If sPilih Then
                .TextMatrix(i, 0) = -1
            Else
                .TextMatrix(i, 0) = 0
            End If
       
    Next
    .Refresh
End With
End Sub
Public Sub oExportFile()
Dim sTempFileBat As String
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
            
            
        sTempFileBat = App.Path & "\exportfilestemp\"
        If Dir(sTempFileBat, vbDirectory) = "" Then
            MkDir sTempFileBat
        End If
        If Dir(sTempFileBat & "sales.txt") <> "" Then
            Kill sTempFileBat & "sales.txt"
        End If
        If Dir(sTempFileBat & "sales_detail1.txt") <> "" Then
            Kill sTempFileBat & "sales_detail1.txt"
        End If
        
        sQuery = "CALL sp_export_file_faktur('" & App.Path & "\exportfilestemp\','sales')"
        sQuery = Replace(sQuery, "\", "/")
        oCon.Execute sQuery
        
        MsgBox "Proses Export Files Complete ", vbInformation, "Proses Export Files"
        oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Diskon"
End Sub



VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form SiswaKeluarFrm 
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
   Begin Crystal.CrystalReport cr1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame frame1 
      BackColor       =   &H8000000A&
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   9420
         TabIndex        =   23
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
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         EmptyButtonCaption=   "None"
      End
      Begin VB.Frame frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Status"
         Height          =   615
         Index           =   2
         Left            =   7320
         TabIndex        =   20
         Top             =   960
         Width           =   3855
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "Close"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000A&
            Caption         =   "Open"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   1335
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
         Index           =   5
         Left            =   2340
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3000
         Width           =   8775
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
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2640
         Width           =   8775
      End
      Begin VB.Frame frame1 
         BackColor       =   &H8000000A&
         Caption         =   "Alasan Keluar"
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   7095
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000A&
            Caption         =   "Lain-lain"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   15
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000A&
            Caption         =   "Pindah Keluar"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   14
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000A&
            Caption         =   "Keluar "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
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
         Height          =   645
         Index           =   3
         Left            =   2340
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   4935
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
         Top             =   240
         Width           =   1695
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
         Left            =   9420
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   11160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaKeluarFrm.frx":0000
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
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "SiswaKeluarFrm.frx":001C
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
         Caption         =   "Alamat"
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Referensi"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   4
         Left            =   7320
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No.ID Siswa"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No Dokumen"
         Height          =   315
         Index           =   0
         Left            =   7320
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "SiswaKeluarFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim snodokumen As String
Dim stgldokumen As String
Dim sdokstatus As String
Dim skodekeluar As String
Dim sketerangan As String
Dim sreferensi As String
Dim snoidsiswa As String
Dim snoidsiswasebelum As String
Dim sobjtype As String
Dim saudituser As String
Dim sauditdate As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "Select * from vtransaksi_siswa_keluar where nodokumen='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52
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
    sQuery = "Select *  from vtransaksi_siswa_keluar order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_siswa_keluar where nodokumen >'" & Text1(0).Text & "' order by nodokumen asc limit 1"
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
    sQuery = "Select  *  from vtransaksi_siswa_keluar where nodokumen<'" & Text1(0).Text & "' order by nodokumen desc limit 1"
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
    sQuery = "Select *  from vtransaksi_siswa_keluar order by nodokumen desc limit 1 "
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
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52
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
        oCon.Execute "Call spUpdate_Status_Siswa('" & snoidsiswasebelum & "','" & snoidsiswa & "','1')"
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
        sQuery = "Delete from transaksi_siswa_keluar where nodokumen='" & snodokumen & "'"
        oCon.Execute sQuery
        oCon.Execute "Call spUpdate_Status_Siswa('" & snoidsiswasebelum & "','" & snoidsiswa & "','1')"
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52
    Text1(0).Locked = False
    Text1(1).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Option1(0).value = True
    Option2(0).value = True
    FlatDatePicker1.value = Now()
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler

    If istatus = StatusForm.DataBaru Then
        If Text1(0) = "" Then
            snodokumen = GetDocnum(transaksi_siswakeluar, True, parkir)
            Text1(0).Text = snodokumen
        Else
            snodokumen = Text1(0).Text
        End If
    Else
        snoidsiswa = Text1(0).Text
    End If
    
    snodokumen = Text1(0).Text
    snoidsiswa = Text1(1).Text
    sdokstatus = IIf(Option2(0).value = True, "1", "0")
    stgldokumen = Format(FlatDatePicker1.value, "yyyy-mm-dd")
    sketerangan = Text1(4).Text
    sreferensi = Text1(5).Text
    skodekeluar = IIf(Option1(0).value = True, "1", IIf(Option1(1).value = True, "2", "3"))
    
    sUpdate = "update transaksi_siswa_keluar set "
    sUpdate = sUpdate & "kodekeluar='" & skodekeluar & "' , "
    sUpdate = sUpdate & "tgldokumen='" & stgldokumen & "' , "
    sUpdate = sUpdate & "noidsiswa='" & snoidsiswa & "' , "
    sUpdate = sUpdate & "dokstatus='" & sdokstatus & "' , "
    sUpdate = sUpdate & "keterangan='" & sketerangan & "' , "
    sUpdate = sUpdate & "referensi='" & sreferensi & "'  "
    'sUpdate = sUpdate & " where noidsiswa"
    sUpdate = sUpdate & " where nodokumen='" & snodokumen & "'"
    
    sInsert = "insert into transaksi_siswa_keluar ("
    sInsert = sInsert & "nodokumen,tgldokumen,dokstatus,kodekeluar,noidsiswa,keterangan,referensi ) values "
    sInsert = sInsert & "("
    sInsert = sInsert & "'" & snodokumen & "',"
    sInsert = sInsert & "'" & stgldokumen & "',"
    sInsert = sInsert & "'" & sdokstatus & "',"
    sInsert = sInsert & "'" & skodekeluar & "',"
    sInsert = sInsert & "'" & snoidsiswa & "',"
    sInsert = sInsert & "'" & sketerangan & "',"
    sInsert = sInsert & "'" & sreferensi & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsSiswaKeluar, ""
    If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsSiswa, "stssiswa='1'"
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
        Text1(3) = oFindByQuery("select CONCAT(b.almtrumah1,IF(b.almtrumah2='','',CONCAT(',',b.almtrumah2))) AS alamat from master_siswa b where noidsiswa='" & Text1(1) & "'", parkir)
    End If
End Select

Set oBrowse = Nothing
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Form Siswa Keluar"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, 52

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
End Sub

Private Sub Form_Load()
oFormatOption 2, Me
snoidsiswasebelum = ""
cleardata
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    snoidsiswasebelum = ToText(oRs("noidsiswa"))
    Text1(0).Text = oRs("nodokumen")
    FlatDatePicker1.value = oRs("tgldokumen")
    KodeUserAksesTemp = oRs("nodokumen")
    Text1(0).Locked = True
    Text1(1).Text = ToText(oRs("noidsiswa"))
    Text1(2).Text = ToText(oRs("nmlengkap"))
    Text1(3).Text = ToText(oRs("alamat"))
    Text1(4).Text = ToText(oRs("keterangan"))
    Text1(5).Text = ToText(oRs("referensi"))
    Select Case ToText(oRs("kodekeluar"))
    Case "1"
        Option1(0).value = True
    Case "2"
        Option1(1).value = True
    Case "3"
        Option1(2).value = True
    End Select
    Option2(0).value = IIf(oRs("dokstatus") = "1", True, False)
    Option2(1).value = IIf(oRs("dokstatus") = "0", True, False)
    'Text1(2).Text = DecryptPassword(oRs("Password")) dokstatus
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
Text1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
'Text1(Index).BackColor = &H80000005
'If Index = 0 Then FindData Text1(0).Text
End Sub

Public Sub Execution()
On Error GoTo errhandler
Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\SiswaKeluarFrm.Rpt"

Dim sKriteria As String

sQuery = "SELECT * from vtransaksi_siswa_keluar_rpt vtransaksi_siswa_keluar_rpt1 where nodokumen='" & Text1(0) & "'"

Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(1) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(2) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(3) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Form Siswa Keluar"
End Sub


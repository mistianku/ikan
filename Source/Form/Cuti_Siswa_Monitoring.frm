VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Cuti_Siswa_Monitoring 
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Daftar Siswa Cuti"
      Height          =   5175
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   10455
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   4815
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10095
         _cx             =   17806
         _cy             =   8493
         Appearance      =   2
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
         BackColorSel    =   8453888
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12566210
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Cuti_Siswa_Monitoring.frx":0000
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
      Caption         =   "Status Dokumen"
      Height          =   735
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Caption         =   "Aktif"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Caption         =   "Tidak Aktif"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Caption         =   "Semua Status"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Konfirmasi"
      Height          =   735
      Index           =   2
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Semua Status"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Sudah Konfirmasi"
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Belum Konfirmasi"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   11160
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Sort By Status Konfirmasi"
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Tidak"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Ya"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Cari Berdasrkan"
      Height          =   1215
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   10455
      Begin VB.OptionButton Option5 
         BackColor       =   &H8000000A&
         Caption         =   "Nama Siswa"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H8000000A&
         Caption         =   "No.ID.Siswa"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   2055
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
         Left            =   120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Cuti"
      Height          =   735
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   10455
      Begin VSDFLATS.FlatComboBox Combo1 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
         MouseIcon       =   "Cuti_Siswa_Monitoring.frx":01BC
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Masuk Kembali"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   8520
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Selesai Cuti"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   7080
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Mulai Cuti"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Dokumen"
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan Dari Tanggal"
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Cuti_Siswa_Monitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim sBatasTanggal As String
Dim sKriteria As String

Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Dim sKriteria As String
Dim sstssiswa As String

If Option1(0).value = True Then
    sKriteria = "stssiswa='1'"
End If
If Option1(1).value = True Then
    sKriteria = "stssiswa='0'"
End If
If Option1(2).value = True Then
    sKriteria = "''=''"
End If

Select Case Index
Case 0
    oBrowse.ShowFinder BrowsSiswa, sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsSiswa, sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Combo1_Change()
ShowGrid1 fQuery
End Sub

Private Sub Combo1_Click()
ShowGrid1 fQuery
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
ShowGrid1 fQuery
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Form Cuti Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringCutiSiswa
End Sub

Private Sub Form_Load()
Dim irow As Integer
For irow = 1 To 12
Combo1.AddItem irow
Next
Combo1.Text = 3
cleardata

For irow = 0 To Option1.Count - 1
    Option1(irow).BackColor = warna(sWarnaBackcolour)
    Option1(irow).ForeColor = warna(sWarnaLabel)
Next
For irow = 0 To Option2.Count - 1
    Option2(irow).BackColor = warna(sWarnaBackcolour)
    Option2(irow).ForeColor = warna(sWarnaLabel)
Next
For irow = 0 To Option3.Count - 1
    Option3(irow).BackColor = warna(sWarnaBackcolour)
    Option3(irow).ForeColor = warna(sWarnaLabel)
Next
For irow = 0 To Option4.Count - 1
    Option4(irow).BackColor = warna(sWarnaBackcolour)
    Option4(irow).ForeColor = warna(sWarnaLabel)
Next
For irow = 0 To Option5.Count - 1
    Option5(irow).BackColor = warna(sWarnaBackcolour)
    Option5(irow).ForeColor = warna(sWarnaLabel)
Next
Combo1.Text = 6
istatus = Normal
cleardata
'BrowseUserID(0).Top = Text1(0).Top
'BrowseUserID(0).Height = Text1(0).Height
'BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width
'
'BrowseUserID(1).Top = Text1(2).Top
'BrowseUserID(1).Height = Text1(2).Height
'BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

'sQuery = fQuery
ShowGrid1 fQuery
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub
Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sstssiwa As String
Dim sdetail As String
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim ssortbykonfirmasists As String
Dim ssortbykonfirmsts As String
Dim sstatusdok As String

If Option4(0).value = True Then
    sstatusdok = "1"
End If
If Option4(1).value = True Then
    sstatusdok = "0"
End If
If Option4(2).value = True Then
    sstatusdok = "2"
End If

If Text1(0).Text = "" Then
    snoidsiswafr = oFindByQuery("select noidsiswa from master_siswa order by noidsiswa asc limit 1 ", parkir)
Else
    snoidsiswafr = Text1(0).Text
End If
'If Text1(2).Text = "" Then
'    snoidsiswato = oFindByQuery("select noidsiswa from master_siswa order by noidsiswa desc limit 1 ", parkir)
'Else
'    snoidsiswato = Text1(2).Text
'End If

Me.CR1.Reset
Me.CR1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.CR1.ReportFileName = App.Path + "\Reports\cuti_rpt.Rpt"

Dim sKriteria As String
If Option1(0).value = True Then
sKriteria = " where konfirmasists  = '0'"
sstssiwa = 1
End If
If Option1(1).value = True Then
sKriteria = " where konfirmasists  = '1'"
sstssiwa = 0
End If
If Option1(2).value = True Then
sKriteria = " where '1'  = '1'"
sstssiwa = 2
End If

sKriteria = sKriteria & " and noidsiswa between '" & snoidsiswafr & "' and '" & snoidsiswato & "'"
If Option2(0).value = True Then
    ssortbykonfirmasists = "1"
Else
    ssortbykonfirmasists = "0"
End If

'stglfr = Format(FlatDatePicker1(0).CurrentDate, "YYYY-MM-DD")
'stglto = Format(FlatDatePicker1(1).CurrentDate, "YYYY-MM-DD")
sKriteria = sKriteria & " and tanggal between '" & stglfr & "' and '" & stglto & "'"
sKriteria = sKriteria & " and dokumensts=" & IIf(sstatusdok = "2", "dokumensts", "'" & sstatusdok & "'")

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vtransaksi_cuti_rpt vtransaksi_cuti_rpt1 " & sKriteria & " "
'
'
Me.CR1.SQLQuery = sQuery
Me.CR1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.CR1.ParameterFields(1) = "sortbykonfirmasists" & ";" & ssortbykonfirmasists & ";" & True
Me.CR1.ParameterFields(2) = "tglfr" & ";" & stglfr & ";" & True
Me.CR1.ParameterFields(3) = "tglto" & ";" & stglto & ";" & True
Me.CR1.ParameterFields(4) = "sortbykonfirmsts" & ";" & sstssiwa & ";" & True
Me.CR1.ParameterFields(5) = "noidsiswafr" & ";" & snoidsiswafr & ";" & True
Me.CR1.ParameterFields(6) = "noidsiswato" & ";" & snoidsiswato & ";" & True
Me.CR1.ParameterFields(7) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.CR1.ParameterFields(8) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.CR1.ParameterFields(9) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True

Me.CR1.Destination = crptToWindow
Me.CR1.RetrieveDataFiles
Me.CR1.WindowState = crptMaximized
Me.CR1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Form Cuti Siswa"
End Sub
Public Sub ShowGrid1(sQuery As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    'sQuery = "    SELECT * from vtransaksi_cuti_rpt   " & sKondisi
    'sQuery = sQuery & "   LEFT JOIN master_default_pelajaran AS b ON a.pelajaran=b.pelajaran "
    'sQuery = sQuery & "   WHERE a.noidsiswa='" & keynoidsiswa & "' and stskelas='1'"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid1

        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail ogrid1
        '.ColHidden(.Cols - 1) = True
        '.Cols = 4
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("nodokumen"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("tanggal"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("noidsiswa"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("nmlengkap"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("almtrumah1"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("tglmulaicuti"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("tglselesaicuti"))
                .TextMatrix(i, 7) = RTrim(oRsDetail("tglmasukkembali")) 'konfirmasistsdesc
                .TextMatrix(i, 8) = RTrim(oRsDetail("konfirmasistsdesc")) 'konfirmasistsdesc
                .TextMatrix(i, 9) = RTrim(oRsDetail("tglkonfirmasi2"))
                .TextMatrix(i, 10) = RTrim(oRsDetail("ketkonfirmasi"))
                .TextMatrix(i, 11) = RTrim(oRsDetail("batastanggal"))
                .TextMatrix(i, 12) = RTrim(oRsDetail("sisa"))
                If ToNumber(oRsDetail("sisa")) < 1 Then
                    .Cell(flexcpForeColor, i, 0, , .Cols - 1) = vbRed
                End If
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
            'ShowGrid2 .TextMatrix(.row, .Cols - 1), ToNumber(fcTahun.Text)
            
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Function fQuery() As String
Dim sstssiwa As String
Dim sdetail As String
Dim snoidsiswafr As String
Dim snoidsiswato As String
Dim stglfr As String
Dim stglto As String
Dim ssortbykonfirmasists As String
Dim ssortbykonfirmsts As String
Dim sstatusdok As String
Dim oBatasTanggal As String

If Option4(0).value = True Then
    sstatusdok = "1"
End If
If Option4(1).value = True Then
    sstatusdok = "0"
End If
If Option4(2).value = True Then
    sstatusdok = "2"
End If

If Option5(0).value = True Then
    snoidsiswafr = "noidsiswa like '%" & ToText(Text1(0)) & "%'"
Else
    snoidsiswafr = "nmlengkap like '%" & ToText(Text1(0)) & "%'"
End If

Dim sKriteria As String
If Option1(0).value = True Then
sKriteria = " where konfirmasists  = '0'"
sstssiwa = 1
End If
If Option1(1).value = True Then
sKriteria = " where konfirmasists  = '1'"
sstssiwa = 0
End If
If Option1(2).value = True Then
sKriteria = " where '1'  = '1' and " & snoidsiswafr
sstssiwa = 2
End If
sKriteria = sKriteria & " and " & snoidsiswafr
sKriteria = sKriteria & "  "
If Option2(0).value = True Then
    ssortbykonfirmasists = "1"
Else
    ssortbykonfirmasists = "0"
End If
sKriteria = sKriteria & " and dokumensts=" & IIf(sstatusdok = "2", "dokumensts", "'" & sstatusdok & "'")

If Option3(0).value = True Then
    sBatasTanggal = "tanggal"
    End If
    If Option3(1).value = True Then
    sBatasTanggal = "tglmulaicuti"
    End If
    If Option3(2).value = True Then
    sBatasTanggal = "tglselesaicuti"
    End If
    If Option3(3).value = True Then
    sBatasTanggal = "tglmasukkembali"
    End If
    'DATE_ADD(  tanggal,INTERVAL 3 MONTH)
    oBatasTanggal = "DATE_ADD(" & sBatasTanggal & ",INTERVAL " & Combo1.Text & " MONTH) as batastanggal,DATEDIFF(DATE_ADD(" & sBatasTanggal & ",INTERVAL " & Combo1.Text & " MONTH) , now()) as sisa "
    
sQuery = "SELECT"
sQuery = sQuery & "    * ," & oBatasTanggal
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vtransaksi_cuti_rpt2 vtransaksi_cuti_rpt1 " & sKriteria & " order by sisa asc"
fQuery = sQuery
End Function

Private Sub oGrid1_DblClick()
Transaksi_CutiFrm2.Show
Transaksi_CutiFrm2.FindData Cuti_Siswa_Monitoring.ogrid1.TextMatrix(Cuti_Siswa_Monitoring.ogrid1.row, 0)
End Sub

Private Sub Option1_Click(Index As Integer)
ShowGrid1 fQuery
End Sub

Private Sub Option3_Click(Index As Integer)
ShowGrid1 fQuery
End Sub

Private Sub Option4_Click(Index As Integer)
ShowGrid1 fQuery
End Sub

Private Sub Option5_Click(Index As Integer)
ShowGrid1 fQuery
End Sub

Private Sub Text1_Change(Index As Integer)
ShowGrid1 fQuery
End Sub

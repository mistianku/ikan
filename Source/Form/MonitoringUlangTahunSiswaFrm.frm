VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MonitoringUlangTahunSiswaFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   5835
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   10995
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
   ScaleWidth      =   10995
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame2"
      Height          =   6015
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   10455
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10215
         _cx             =   18018
         _cy             =   9975
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   8454143
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"MonitoringUlangTahunSiswaFrm.frx":0000
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
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Tanggal Ulang Tahun"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   4
         Top             =   360
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
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   5
         Top             =   720
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
      Begin VB.Label Label1 
         Caption         =   "Sampai"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Siswa"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   10455
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tidak Aktif"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cuti"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Aktif"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   8640
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "MonitoringUlangTahunSiswaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim scek1 As Integer
Dim scek2 As Integer
Dim scek3 As Integer



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Dim sKriteria As String
Dim sstssiswa As String



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

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    scek1 = Check1(Index).value
Case 1
    scek2 = Check1(Index).value
Case 2
    scek3 = Check1(Index).value
End Select
ShowGridDaftarSiswaUlangTahun Day(FlatDatePicker1(0).value), Month(FlatDatePicker1(0).value), Day(FlatDatePicker1(1).value), Month(FlatDatePicker1(0).value), scek1, scek2, scek3
End Sub

Private Sub FlatDatePicker1_Change(Index As Integer)
ShowGridDaftarSiswaUlangTahun Day(FlatDatePicker1(0).value), Month(FlatDatePicker1(0).value), Day(FlatDatePicker1(1).value), Month(FlatDatePicker1(0).value), scek1, scek2, scek3
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Monitoring Siswa Ulang Tahun"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringUlangTahunSiswa
End Sub

Private Sub Form_Load()
'oFormatOption 2, Me
oFormatCheckList 1, Me
istatus = Normal
cleardata

    scek1 = Check1(0).value
    scek2 = Check1(1).value
    scek3 = Check1(2).value

FlatDatePicker1(0).value = Now()
FlatDatePicker1(1).value = Now()
ShowGridDaftarSiswaUlangTahun Day(FlatDatePicker1(0).value), Month(FlatDatePicker1(0).value), Day(FlatDatePicker1(1).value), Month(FlatDatePicker1(0).value), scek1, scek2, scek3
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

stglfr = Format(FlatDatePicker1(0).value, "YYYY-MM-DD")
stglto = Format(FlatDatePicker1(1).value, "YYYY-MM-DD")


Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\DaftarSiswaUlangTahun.Rpt"

Dim sKriteria As String
'WHERE MONTH(`ns`.`tgllahir`)*100+DAY(`ns`.`tgllahir`) BETWEEN sbulanfr*100+stanggalfr AND sbulanto*100+stanggalto
'AND (`ns`.`stssiswa`=IF(scek1='1','1','9') OR `ns`.`stssiswa`=IF(scek2='1','2','9') OR `ns`.`stssiswa`=IF(scek3='1','0','9'))
sKriteria = " Where MONTH(tgllahir)*100+DAY(tgllahir) BETWEEN " & Month(stglfr) * 100 + Day(stglfr) & " and " & Month(stglto) * 100 + Day(stglto)
sKriteria = sKriteria & " AND (`stssiswa`=IF(" & scek1 & "=1,'Aktif','9') OR `stssiswa`=IF(" & scek2 & "=1,'Cuti','9') OR `stssiswa`=IF(" & scek3 & "=1,'Tidak Aktif','9'))"



sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vdaftar_ulangtahun vdaftar_ulangtahun1 " & sKriteria & " order by nlahir "
'
'
Me.cr1.SQLQuery = sQuery

Me.cr1.ParameterFields(0) = "cmpnyname" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(1) = "address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
Me.cr1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(4) = "tglfr" & ";" & Format(stglfr, "DD-MM-YYYY") & ";" & True
Me.cr1.ParameterFields(5) = "tglto" & ";" & Format(stglto, "DD-MM-YYYY") & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Public Sub ShowGridDaftarSiswaUlangTahun(stglfr As Integer, sblnfr As Integer, stglto As Integer, sblnto As Integer, scek1 As Integer, scek2 As Integer, scek3 As Integer)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim sQuery As String
    Dim sSiswaCount As Integer
    Dim sTotalBiaya As Double
    sQuery = "CALL sp_get_siswa_ulang_tahun(" & stglfr & "," & sblnfr & "," & stglto & "," & sblnto & "," & scek1 & "," & scek2 & "," & scek3 & ")"
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

        GridModul.ClearGridDetail oGrid1

        If Not oRsDetail.EOF Then
            Dim i As Double
            
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("noidsiswa"))
                .TextMatrix(i, 1) = ToText(oRsDetail("nmlengkap"))
                .TextMatrix(i, 2) = ToText(oRsDetail("jnskelamin"))
                .TextMatrix(i, 3) = ToText(oRsDetail("tptlahir"))
                .TextMatrix(i, 4) = ToText(oRsDetail("tgllahir"))
                .TextMatrix(i, 5) = (oRsDetail("almtrumah1"))
                .TextMatrix(i, 6) = ToText(oRsDetail("notelprumah"))
                .TextMatrix(i, 7) = ToText(oRsDetail("stssiswa"))
                .TextMatrix(i, 8) = ToText(oRsDetail("kelas"))
'                If .TextMatrix(i, 3) = "Cuti" Then
'                .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
'                .Cell(flexcpForeColor, 1, 0, , .Cols - 1) = vbRed
'                End If
                oRsDetail.MoveNext
            Loop
            '.Select 1, 0
            'ShowGrid2 .TextMatrix(.row, .Cols - 1), ToNumber(fcTahun.Text)
            
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

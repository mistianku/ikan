VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form MonitoringPembayaran 
   BackColor       =   &H8000000A&
   Caption         =   "Master Data User"
   ClientHeight    =   9120
   ClientLeft      =   -135
   ClientTop       =   645
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
   ScaleHeight     =   9120
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Status Pembayaran"
      Height          =   735
      Index           =   3
      Left            =   8280
      TabIndex        =   22
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton Option1 
         Caption         =   "Belum"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sudah"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cuti"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   12000
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Kursus"
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   1455
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   2566
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
         ForeColorSel    =   -2147483641
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
         FormatString    =   $"MonitoringPembayaran.frx":0000
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
      Caption         =   "Daftar Siswa"
      Height          =   5655
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   11775
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   9840
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   360
         Width           =   9375
      End
      Begin VSFlex8LCtl.VSFlexGrid ogrid1 
         Height          =   4215
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   11535
         _cx             =   20346
         _cy             =   7435
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
         ForeColorSel    =   -2147483641
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
         AllowUserResizing=   0
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
         FormatString    =   $"MonitoringPembayaran.frx":00C0
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
         Caption         =   "Total"
         Height          =   315
         Index           =   3
         Left            =   7680
         TabIndex        =   9
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Siswa"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama "
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Periode Pembayaran"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         Border          =   0   'False
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
         MouseIcon       =   "MonitoringPembayaran.frx":0166
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   3
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         Border          =   0   'False
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
         MouseIcon       =   "MonitoringPembayaran.frx":0182
      End
      Begin VB.Label Label1 
         Caption         =   "Periode"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Berdasarkan Periode"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.OptionButton Option2 
         Caption         =   "Tanggal"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bulan"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tanggal Pembayaran"
      Height          =   735
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   8055
      Begin NeoCalendarII.DatePicker DatePicker1 
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   20
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
      Begin NeoCalendarII.DatePicker DatePicker1 
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   21
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
      Begin VB.Label Label1 
         Caption         =   "S/D Tanggal"
         Height          =   315
         Index           =   5
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Tanggal"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "MonitoringPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim sproses1 As Integer
Dim sproses2 As Integer
Dim sproses3 As Integer
Dim snoidsiswa As String
Dim sjnsbayar As String
Dim stsbayar As String

Private Sub BrowseUserID_Click(Index As Integer)
    Dim oBrowse As New BrowseFrm
Select Case Index
Case 0

    oBrowse.ShowFinder BrowsBrand, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsBrand, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub


Private Sub FlatComboBox1_Click(Index As Integer)
ShowGridDaftarPembayaran ToNumber(FlatComboBox1(0).Text), FlatComboBox1(1).ListIndex + 1, Trim(Text1(0)), sjnsbayar
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Monitoring Pembayaran Siswa"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMonitoringPembayaranSiswa
End Sub

Private Sub Form_Load()
oFormatFrameBackground Frame1(0)
oFormatOption 2, Me
istatus = Normal
cleardata
oAddComboBoxTahun
oGetComboBulanan
sjnsbayar = 0
MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"
ShowGridDaftarPembayaran ToNumber(FlatComboBox1(0).Text), FlatComboBox1(1).ListIndex + 1, Trim(Text1(0)), sjnsbayar
Frame1(2).Top = Frame1(0).Top
If Option2(0).value = True Then
    Frame1(0).Visible = True
    Frame1(2).Visible = False
Else
    Frame1(0).Visible = False
    Frame1(2).Visible = True
End If

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
Public Sub oAddComboBoxTahun()
Dim i As Integer
Dim sawal As Integer
sawal = ToNumber(oFindByQuery("select min(year(tglpendaftaran)) from transaksi_pendaftaran limit 1", parkir))
For i = IIf(sawal > Year(Now), Year(Now), sawal) To Year(Now)
    FlatComboBox1(0).AddItem i
Next
FlatComboBox1(0).Text = Year(Now())
End Sub


Public Sub oGetComboBulanan()
On Error GoTo errhandler
Dim i As Integer

        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.parkir)
        sQuery = "select namabulan from master_bulan"
        Set oRs = oCon.Execute(sQuery) 'master_moduleaccess
        With oRs
        Do While Not .EOF
            FlatComboBox1(1).AddItem .Fields(0)
        .MoveNext
        Loop
        FlatComboBox1(1).ListIndex = Month(Date) - 1
        End With
        oCon.Close
        
        istatus = Normal
        Exit Sub

errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Sub
Public Sub ShowGridDaftarPembayaran(syop As Integer, smop As Integer, sNamaLengkap As String, sStatusBayar As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim sQuery As String
    Dim sSiswaCount As Integer
    Dim sTotalBiaya As Double
    sQuery = "CALL sp_get_monitoring_bayar(" & syop & "," & smop & ",'" & sNamaLengkap & "','" & sStatusBayar & "')"
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(0)

        GridModul.ClearGridDetail oGrid1(0)
        GridModul.ClearGridDetail oGrid1(1)
        If Not oRsDetail.EOF Then
            Dim i As Double
            
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("Nourt"))
                .TextMatrix(i, 1) = ToText(oRsDetail("noidsiswa"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("nmlengkap"))
                .TextMatrix(i, 3) = ToText(oRsDetail("alamat"))
                .TextMatrix(i, 4) = ToNumber(oRsDetail("biaya"))
                
                sSiswaCount = sSiswaCount + 1
                sTotalBiaya = sTotalBiaya + ToNumber(oRsDetail("biaya"))
                
                oRsDetail.MoveNext
            Loop
            .Select 1, 0
            Text1(1) = formatRupiah(ToNumber(sSiswaCount))
            Text1(2) = formatRupiah(ToNumber(sTotalBiaya))
            'ShowGrid2 .TextMatrix(.row, .Cols - 1), ToNumber(fcTahun.Text)
            
            ShowGridDaftarKelas .TextMatrix(.row, 1)
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Private Sub oGrid1_Click(Index As Integer)
With oGrid1(0)
If .row = 0 Then Exit Sub
    ShowGridDaftarKelas .TextMatrix(.row, 1)
End With
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    sjnsbayar = "0"
Case 1
    sjnsbayar = "3"
Case 2
    sjnsbayar = "2"
End Select
ShowGridDaftarPembayaran ToNumber(FlatComboBox1(0).Text), FlatComboBox1(1).ListIndex + 1, Trim(Text1(0)), sjnsbayar
End Sub

Private Sub Option2_Click(Index As Integer)
If Option2(0).value = True Then
    Frame1(0).Visible = True
    Frame1(2).Visible = False
Else
    Frame1(0).Visible = False
    Frame1(2).Visible = True
End If
End Sub

Private Sub Text1_Change(Index As Integer)
If Index = 0 Then
    ShowGridDaftarPembayaran ToNumber(FlatComboBox1(0).Text), FlatComboBox1(1).ListIndex + 1, Trim(Text1(0)), sjnsbayar
End If
End Sub
Public Sub ShowGridDaftarKelas(snoidsiswa As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
    Dim sQuery As String
    Dim sSiswaCount As Integer
    Dim sTotalBiaya As Double
    sQuery = "CALL sp_get_monitoring_kelas_info('" & snoidsiswa & "')"
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1(1)

        GridModul.ClearGridDetail oGrid1(1)

        If Not oRsDetail.EOF Then
            Dim i As Double
            
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = ToText(oRsDetail("nourt"))
                .TextMatrix(i, 1) = ToText(oRsDetail("nokursus"))
                .TextMatrix(i, 2) = (oRsDetail("tglmulai"))
                .TextMatrix(i, 3) = ToText(oRsDetail("statuskls"))
                .TextMatrix(i, 4) = ToText(oRsDetail("kelas"))
                .TextMatrix(i, 5) = ToText(oRsDetail("tingkatansek"))
                If .TextMatrix(i, 3) = "Cuti" Then
                .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
                .Cell(flexcpForeColor, 1, 0, , .Cols - 1) = vbRed
                End If
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

Public Sub Execution()
On Error GoTo errhandler
Dim syop As Integer
Dim smop As String
Dim snama As String
Dim sdetail As String

If MsgBox("Format Detail !!", vbYesNo) = vbYes Then
    sdetail = "Y"
Else
    sdetail = "N"
End If

syop = (FlatComboBox1(0).Text)
smop = IIf(Len(Trim(FlatComboBox1(1).ListIndex + 1)) = 1, "0" & FlatComboBox1(1).ListIndex + 1, FlatComboBox1(1).ListIndex + 1)
snama = "%" & Text1(0) & "%"
Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\monitoring_pembayaran_rpt.rpt"

Dim sKriteria As String

sKriteria = " Where yop=" & syop & " and mop=" & smop & " and nmlengkap like '" & snama & "' and "
sKriteria = sKriteria & " jnsbayar ='" & sjnsbayar & "'"
sQuery = "Select * from vget_monitoring_pembayaran_rpt  vget_monitoring_pembayaran_rpt1 "
sQuery = sQuery & sKriteria

Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
Me.cr1.ParameterFields(1) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
Me.cr1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
Me.cr1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
Me.cr1.ParameterFields(4) = "yop" & ";" & syop & ";" & True
Me.cr1.ParameterFields(5) = "mop" & ";" & smop & ";" & True
Me.cr1.ParameterFields(6) = "jnsbayar" & ";" & sjnsbayar & ";" & True
Me.cr1.ParameterFields(7) = "detail" & ";" & sdetail & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

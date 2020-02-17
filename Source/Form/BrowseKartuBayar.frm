VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form BrowseKartuBayar 
   Appearance      =   0  'Flat
   Caption         =   "Kartu Administrasi Bulanan"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   10455
   StartUpPosition =   1  'CenterOwner
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   6960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      MouseIcon       =   "BrowseKartuBayar.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ambil"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kartu Pembayaran"
      Height          =   3615
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   9975
      Begin VSFlex8LCtl.VSFlexGrid oGrid2 
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   9615
         _cx             =   16960
         _cy             =   5741
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         BackColorAlternate=   -2147483643
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
         Rows            =   13
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"BrowseKartuBayar.frx":001C
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
      Caption         =   "Kelas"
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   9975
      Begin VSFlex8LCtl.VSFlexGrid oGrid1 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9615
         _cx             =   16960
         _cy             =   1931
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         FormatString    =   $"BrowseKartuBayar.frx":00A2
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
      Caption         =   "Siswa"
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.OptionButton Option1 
         Caption         =   "Diatas SMA"
         Height          =   375
         Index           =   4
         Left            =   7800
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SMA"
         Height          =   375
         Index           =   3
         Left            =   6960
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SMP"
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SD"
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "TK"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   11
         Top             =   960
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   6975
      End
      Begin VSDFLATS.FlatComboBox fcTahun 
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "BrowseKartuBayar.frx":0148
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Periode Tahun"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Nama Siswa"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "No.ID"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VSDFLATS.FlatButton Command1 
      Height          =   375
      Index           =   1
      Left            =   5160
      TabIndex        =   17
      Top             =   6960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      MouseIcon       =   "BrowseKartuBayar.frx":0164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Batal"
   End
End
Attribute VB_Name = "BrowseKartuBayar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim snilaibayar As Double
Dim stingkatansklh As String
Dim stxtbiaya As String
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    oAmbilKartuBayar KwitansiFrm.lbldocentry, KwitansiFrm.lbllinenum
Case 1
    Unload Me
End Select
End Sub

Private Sub fcTahun_Click()
'ShowGrid2 .TextMatrix(.row, .Cols - 1), fcTahun.Text
With oGrid1
    ShowGrid2 .TextMatrix(.row, .Cols - 1), fcTahun.Text
End With
End Sub

Private Sub Form_Load()
oClearData
GridModul.ClearGridDetail oGrid2
oFormatWarnaLabel merahtua, hijaubiasa, background, Me
oFormatFrameBackground Frame1(0)
'oFormatFrameBackground Frame2
'oFormatFrameBackground Frame3
Me.BackColor = warna(background)
Dim i As Integer
For i = 0 To Option1.Count - 1
    Option1(i).BackColor = warna(background)
Next

Text1(0).Text = KwitansiFrm.Text1(4).Text
Text1(1).Text = KwitansiFrm.Text1(5).Text
stingkatansklh = oFindByQuery("select tingkatansklh from master_siswa where noidsiswa='" & Text1(0).Text & "'", DBaseConection.Modul)
oAddfcTahun
Option1(ToNumber(stingkatansklh)).value = True
Select Case ToNumber(stingkatansklh)
Case 0
    stxtbiaya = "biayatk"
Case 1
    stxtbiaya = "biayasd"
Case 2
    stxtbiaya = "biayasmp"
Case 3
    stxtbiaya = "biayasma"
Case 4
    stxtbiaya = "biayadiatassma"
End Select

snilaibayar = oFindByQuery("select " & stxtbiaya & " from master_default_biaya where biayaid=2", DBaseConection.Modul)
With oGrid1
    ShowGrid2 .TextMatrix(.row, .Cols - 1), fcTahun.Text
End With
End Sub

Public Sub oClearData()
Dim iText As Integer
For iText = 1 To Text1.Count - 1
    Text1(iText) = ""
Next
        GridModul.ClearGridDetail oGrid1
        GridModul.ClearGridDetail oGrid2
        
End Sub
Public Sub ShowGrid1(keynoidsiswa As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sCekDataDetail1 As Integer
     
    sQuery = "    SELECT a.docentry,a.nokursus,a.tglmulai,ifnull(b.keterangan,'-') AS kelas ,a.noidsiswa,if(a.stskelas='1','Aktif',if(a.stskelas='2','Cuti','Tutup')) as stskelasdesc  FROM master_kelas AS a "
    sQuery = sQuery & "   LEFT JOIN master_default_pelajaran AS b ON a.pelajaran=b.pelajaran "
    sQuery = sQuery & "   WHERE a.noidsiswa='" & keynoidsiswa & "' and stskelas='1'"

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid1

        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)
        GridModul.ClearGridDetail oGrid1
        
        .ColHidden(.Cols - 1) = True
        .Cols = 5
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
                .TextMatrix(i, 3) = RTrim(oRsDetail("stskelasdesc"))
                .TextMatrix(i, .Cols - 1) = RTrim(oRsDetail("docentry"))
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
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
    If oGrid1.Rows = 1 Then Exit Sub
    sQuery = "    SELECT * FROM vmaster_kartu_bayar "
    sQuery = sQuery & "   WHERE docentry = " & keyDocentry & " AND yop = " & keyYop
    sQuery = sQuery & "   Order by mop asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid2
        '.Cols = 6
        '.COLWIDTH(2) = .Width - (.COLWIDTH(0) + .COLWIDTH(1) + .COLWIDTH(5) + 100)

        GridModul.ClearGridDetail oGrid2
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = ToText(oRsDetail("namabulan"))
                Select Case ToText(oRsDetail("jnsbayar"))
                Case "1", "2"
                    .TextMatrix(i, 2) = 0
                Case Else
                    .TextMatrix(i, 2) = snilaibayar
                End Select
                                
                
                .TextMatrix(i, .Cols - 1) = ToText(oRsDetail("stsbayar"))
                If .TextMatrix(i, .Cols - 1) = 0 Then
                    .TextMatrix(i, 0) = -1
                End If
                oRsDetail.MoveNext
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                '1 bypas 2 cuti 3 bayar 0 belum bayar
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

Public Sub oAddfcTahun()
On Error GoTo errhandler
Dim i As Integer
        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "SELECT  MIN(yop) AS yop1,MAX(yop) AS yop2 FROM master_kartu_bayar"
        Set oRs = oCon.Execute(sQuery)
        If Not oRs.EOF Then
            For i = oRs(0) To Year(Now) + 3
                BrowseKartuBayar.fcTahun.AddItem i
            Next
        Else
            fcTahun.AddItem Year(Now)
        End If
        fcTahun.Text = Year(Now)
        oCon.Close
        Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Add Combo Box Tahun"
End Sub

Private Sub oGrid1_Click()
With oGrid1
    ShowGrid2 .TextMatrix(.row, .Cols - 1), fcTahun.Text
End With
End Sub

Private Sub oGrid2_Click()
With oGrid2
Select Case .col
Case 0
    If .TextMatrix(.row, .Cols - 1) = "1" Then
        .Select .row, 0
        .EditCell
        If .TextMatrix(.row, 0) = -1 Then
            .Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
        Else
            .Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbNormal
        End If
    Else
        
    End If
End Select
End With
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    stxtbiaya = "biayatk"
Case 1
    stxtbiaya = "biayasd"
Case 2
    stxtbiaya = "biayasmp"
Case 3
    stxtbiaya = "biayasma"
Case 4
    stxtbiaya = "biayadiatassma"
End Select
snilaibayar = oFindByQuery("select " & stxtbiaya & " from master_default_biaya where biayaid=2", DBaseConection.Modul)
With oGrid1
    ShowGrid2 .TextMatrix(.row, .Cols - 1), fcTahun.Text
End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 0
    ShowGrid1 Text1(0)
Case 1
End Select
End Sub

Public Sub oAmbilKartuBayar(keyDocentry As Integer, keylinenum As Integer)
Dim sdocentry As Integer
Dim sbaselinenum As Integer
Dim slinenum As Integer
Dim sutkbayar As String
Dim sketerangan As String
Dim sbiaya As Double
Dim spotongan As Double
Dim sjumlah As Double
Dim saudituser As String
Dim sauditdate As String

Dim sKelas As String
Dim sTahun As Integer
sketerangan = "Bia.Admin Bulanan"
spotongan = 0
    Dim irow As Integer
    With oGrid2
        For irow = 1 To .Rows - 1
            If .TextMatrix(irow, 0) = -1 And .TextMatrix(irow, .Cols - 1) = "1" Then
                slinenum = slinenum + 1
                sbiaya = sbiaya + ToNumber(.TextMatrix(irow, .Cols - 2))
                sketerangan = sketerangan & " " & Trim(.TextMatrix(irow, 1))
                KwitansiFrm.oGrid2.Rows = KwitansiFrm.oGrid2.Rows + 1
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 0) = oGrid1.TextMatrix(oGrid1.row, 0)
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 1) = fcTahun.Text
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 2) = irow
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 3) = ToNumber(.TextMatrix(irow, .Cols - 2))
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 4) = spotongan
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 5) = ToNumber(.TextMatrix(irow, .Cols - 2)) - spotongan
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 6) = keyDocentry
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 7) = keylinenum
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 8) = slinenum
                KwitansiFrm.oGrid2.TextMatrix(KwitansiFrm.oGrid2.Rows - 1, 9) = 1
                sjumlah = sbiaya - spotongan
            End If
        Next
        
    End With
    If sjumlah = 0 Then
        MsgBox "Tidak Ada Periode Pembayaran Yang dipilih", vbInformation
        BrowseKartuBayar.Show
    Else
        With KwitansiFrm.ogrid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = 2
            .TextMatrix(.Rows - 1, 1) = sketerangan & " No.Kursus:" & oGrid1.TextMatrix(oGrid1.row, 0)
            .TextMatrix(.Rows - 1, 2) = sbiaya
            .TextMatrix(.Rows - 1, 3) = spotongan
            .TextMatrix(.Rows - 1, 4) = sjumlah
            .TextMatrix(.Rows - 1, 5) = keyDocentry
            .TextMatrix(.Rows - 1, 6) = keylinenum
            .TextMatrix(.Rows - 1, 7) = 1
        End With
        Unload Me
    End If
    
End Sub



VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form GLTransRABRkpRpt 
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
   Begin VB.Frame Frame1 
      Caption         =   "Tanggal"
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   10455
      Begin NeoCalendarII.DatePicker DatePicker1 
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
         BorderStyle     =   0
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
         Left            =   2160
         TabIndex        =   22
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
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
         BorderStyle     =   0
         FlatButton      =   0   'False
         AllowEmpty      =   0   'False
         ShowFocusRect   =   0   'False
         UseFocusColor   =   0   'False
         CalendarHeaderForeColor=   -2147483630
         CalendarPresentDateColor=   -2147483646
      End
      Begin VB.Label Label1 
         Caption         =   "Sampai"
         Height          =   315
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari "
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tampil"
      Height          =   855
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "Rekap"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rinci"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   5280
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periode"
      Height          =   1095
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   10455
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   10
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
         MouseIcon       =   "GLTransRABRkpRpt.frx":0000
      End
      Begin VSDFLATS.FlatComboBox FlatComboBox1 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   600
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
         MouseIcon       =   "GLTransRABRkpRpt.frx":001C
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tahun"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1560
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "No. RAB"
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   10455
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
         Left            =   4080
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   720
         Width           =   6135
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
         Left            =   4080
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   360
         Width           =   6135
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
         Left            =   2220
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
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
         Index           =   0
         Left            =   2220
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLTransRABRkpRpt.frx":0038
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
         Left            =   3600
         TabIndex        =   8
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLTransRABRkpRpt.frx":0054
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
         Caption         =   "Sampai"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari "
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GLTransRABRkpRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim syop As Integer
Dim smop As Integer
Dim skodefr As String
Dim skodeto As String
Dim stanggalfr As String
Dim stanggalto As String
Dim sstatus As String
Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0

    oBrowse.ShowFinder BrowsAkunTransRAB, "tanggal between '" & Format(DatePicker1(0).value, "YYYY-MM-DD") & "' and '" & Format(DatePicker1(1).value, "YYYY-MM-DD") & "'"
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsAkunTransRAB, "tanggal between '" & Format(DatePicker1(0).value, "YYYY-MM-DD") & "' and '" & Format(DatePicker1(1).value, "YYYY-MM-DD") & "'"
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If

End Select
    Set oBrowse = Nothing
End Sub

Private Sub Command1_Click()
Execution
End Sub

Private Sub FlatComboBox1_Click(Index As Integer)
syop = FlatComboBox1(0).Text
smop = FlatComboBox1(1).ListIndex + 1
oGetTanggal
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Transaksi RAB Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnTransaksiRABReportRpt
End Sub

Private Sub Form_Load()

DatePicker1(0).value = DateSerial(Year(Now()), Month(Now()), 1)
DatePicker1(1).value = Now()
'the do your printing e.g

'DataReport1.PrintReport

oFormatFrameBackground Frame1(0)
oFormatOption 1, Me
istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"

Dim iTahun As Integer
For iTahun = 2011 To Year(Now()) + 5
FlatComboBox1(0).AddItem iTahun
Next
FlatComboBox1(0).Text = Year(Now())
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Februari"
FlatComboBox1(1).AddItem "Maret"
FlatComboBox1(1).AddItem "April"
FlatComboBox1(1).AddItem "Mei"
FlatComboBox1(1).AddItem "Juni"
FlatComboBox1(1).AddItem "Juli"
FlatComboBox1(1).AddItem "Agustus"
FlatComboBox1(1).AddItem "September"
FlatComboBox1(1).AddItem "Oktober"
FlatComboBox1(1).AddItem "November"
FlatComboBox1(1).AddItem "Desember"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).AddItem "Januari"
FlatComboBox1(1).ListIndex = Month(Now()) - 1
If Option1(0).value = True Then
    sstatus = "1"
Else
    sstatus = "2"
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
Public Sub Execution()
On Error GoTo errhandler
Dim sKriteria As String


skodefr = IIf(Text1(0) = "", oFindByQuery("select min(noslip) from trnent_rab", parkir), Text1(0))
skodeto = IIf(Text1(1) = "", oFindByQuery("select max(noslip) from trnent_rab", parkir), Text1(1))
stanggalfr = Format(DatePicker1(0).value, "YYYY-MM-DD")
stanggalto = Format(DatePicker1(1).value, "YYYY-MM-DD")
'sstatus = IIf(Option1(0).value = True, "Status='Y'", IIf(Option1(1).value = True, "Status='N'", "true"))

If sstatus = "1" Then
        'sKriteria = " where a.noslip between '" & skodefr & "'  and '" & skodeto & "' "
        sKriteria = "and  a.tanggal between '" & stanggalfr & "'  and '" & stanggalto & "' "
        sKriteria = sKriteria & " and a.yop=" & syop & "  and mop =" & smop
        sQuery = "   Select yop*100+mop AS periodeku,tanggal,yop,mop,noslip,a.referensi,a.keterangan,jumlahentri,jumlahtotal,a.closed,   "
        sQuery = sQuery & " linenum,b.kd_rab,c.nm_rab,jumlah,coa_debet,coa_kredit,b.referensi referensi2,b.keterangan keterangan2,CONCAT(yop,'-',mop) AS txtperiode,IF(a.closed='Y','Close','Open') AS txtstatus    "
        sQuery = sQuery & " FROM trnent_rab a INNER JOIN trnent_rabdetail1 b ON a.docentry=b.docentry   " & sKriteria
        sQuery = sQuery & " INNER JOIN tblrab c ON b.kd_rab=c.kd_rab ORDER BY yop*100+mop,tanggal,noslip;   "
        
        With arGLTransRABRkp
            .lblCompany1 = MenuFrm.txtHeader(0)
            .lblCompany2 = MenuFrm.txtHeader(1)
            .lblCompany3 = MenuFrm.txtHeader(2)
            .Label24.Caption = "Transaksi RAB Detail"
            .lblPeriode.Caption = "Periode : " & syop & " -  " & smop
            '.lblPeriode2.Visible = False
            .lblPeriode2.Caption = "Tanggal : " & stanggalfr & " s/d  " & stanggalto
            
            
            '.lblPesan = stxtpesan
            .adoKu.Provider = "MSDASQL.1"
            .adoKu.DataSourceName = MenuFrm.Serverku '"kumonku"
            .adoKu.Source = sQuery
            
            '.PageSettings.Orientation = ddOPortrait
        '    .PageSettings.PaperHeight = MenuFrm.stinggi
        '    .PageSettings.PaperWidth = MenuFrm.slebar
            .Show
            If Not .adoKu.Recordset.EOF() Then
        '    .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
        '    .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value
            End If
        End With

Else

        'sKriteria = " where a.noslip between '" & skodefr & "'  and '" & skodeto & "' "
        sKriteria = "and  a.tanggal between '" & stanggalfr & "'  and '" & stanggalto & "' "
        sKriteria = sKriteria & " and a.yop=" & syop & "  and mop =" & smop
        sQuery = "   Select yop*100+mop AS periodeku,tanggal,yop,mop,noslip,a.referensi,a.keterangan,jumlahentri,jumlahtotal,a.closed,   "
        sQuery = sQuery & " linenum,b.kd_rab,c.nm_rab,jumlah,coa_debet,coa_kredit,b.referensi referensi2,b.keterangan keterangan2,CONCAT(yop,'-',mop) AS txtperiode,IF(a.closed='Y','Close','Open') AS txtstatus    "
        sQuery = sQuery & " FROM trnent_rab a INNER JOIN trnent_rabdetail1 b ON a.docentry=b.docentry   " & sKriteria
        sQuery = sQuery & " INNER JOIN tblrab c ON b.kd_rab=c.kd_rab ORDER BY yop*100+mop,tanggal,noslip;   "
        
        With arGLTransRABRkp2
            .lblCompany1 = MenuFrm.txtHeader(0)
            .lblCompany2 = MenuFrm.txtHeader(1)
            .lblCompany3 = MenuFrm.txtHeader(2)
            .Label24.Caption = "Transaksi RAB Rekap"
            .lblPeriode.Caption = "Periode : " & syop & " -  " & smop
            '.lblPeriode2.Visible = False
            .lblPeriode2.Caption = "Tanggal : " & stanggalfr & " s/d  " & stanggalto
            
            
            '.lblPesan = stxtpesan
            .adoKu.Provider = "MSDASQL.1"
            .adoKu.DataSourceName = MenuFrm.Serverku '"kumonku"
            .adoKu.Source = sQuery
            
            '.PageSettings.Orientation = ddOPortrait
        '    .PageSettings.PaperHeight = MenuFrm.stinggi
        '    .PageSettings.PaperWidth = MenuFrm.slebar
            .Show
            If Not .adoKu.Recordset.EOF() Then
        '    .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
        '    .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value
            End If
        End With


End If


'Me.CR1.ParameterFields(0) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
'Me.CR1.ParameterFields(1) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
'Me.CR1.ParameterFields(2) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
'Me.CR1.ParameterFields(3) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
'Me.CR1.ParameterFields(4) = "yop" & ";" & syop & ";" & True
'Me.CR1.ParameterFields(5) = "mop" & ";" & sbulan & ";" & True

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


Public Sub oGetTanggal()
DatePicker1(0).value = oFindByQuery("select min(tanggal) from trnent_rab where yop=" & syop & " and mop=" & smop, parkir)
DatePicker1(1).value = oFindByQuery("select max(tanggal) from trnent_rab where yop=" & syop & " and mop=" & smop, parkir)
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).value = True Then
    sstatus = "1"
End If
If Option1(1).value = True Then
    sstatus = "2"
End If
End Sub

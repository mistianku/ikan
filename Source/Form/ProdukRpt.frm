VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form ProdukRpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Produk Report"
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
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Report "
      Height          =   1215
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   3000
      Width           =   10455
      Begin VB.OptionButton Option1 
         Caption         =   "Diskon Produk "
         Height          =   375
         Index           =   2
         Left            =   3960
         TabIndex        =   21
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Harga Produk "
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   20
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Produk Info"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Produk"
      Height          =   1335
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1560
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
         Index           =   7
         Left            =   2220
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   360
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
         Index           =   6
         Left            =   2220
         TabIndex        =   12
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
         Index           =   5
         Left            =   4080
         TabIndex        =   11
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
         Index           =   4
         Left            =   4080
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   720
         Width           =   6135
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   14
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukRpt.frx":0000
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
         Index           =   3
         Left            =   3600
         TabIndex        =   15
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProdukRpt.frx":001C
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
         Caption         =   "Dari Kode"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "S/D Kode"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Brand"
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   120
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
         MouseIcon       =   "ProdukRpt.frx":0038
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
         MouseIcon       =   "ProdukRpt.frx":0054
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
         Caption         =   "S/D Kode"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Kode"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "ProdukRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm
Dim sreport As String
Dim skodebrandfr As String
Dim skodebrandto As String
Dim skodeprodukfr As String
Dim skodeprodukto As String

Private Sub BrowseUserID_Click(Index As Integer)
    Dim oBrowse As New BrowseFrm
Select Case Index
Case 0

    oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(2) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder BrowsBrand, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(1) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
Case 2

    oBrowse.ShowFinder BrowsMasterProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(7) = oBrowse.YangDipilih
        Text1(5) = oBrowse.Keterangan
    End If
Case 3
    oBrowse.ShowFinder BrowsMasterProduk, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        Text1(6) = oBrowse.YangDipilih
        Text1(4) = oBrowse.Keterangan
    End If
End Select
    Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Produk Report"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProdukRpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatFrameBackground Frame1(0)
oFormatOption 1, Me
istatus = RefrshRpt
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(1).Top
BrowseUserID(1).Height = Text1(1).Height
BrowseUserID(1).Left = Text1(1).Left + Text1(1).Width
MenuFrm.LblPesanku = "Kode Brand Kosong Berarti Pilih Seluruh Brand"
If Option1(0).value = True Then
    sreport = 1
End If
If Option1(1).value = True Then
    sreport = 2
End If
If Option1(2).value = True Then
    sreport = 3
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
    Text1(i).text = ""
Next
'    Text1(0).Enabled = False
'    Text1(1).Enabled = False
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sKriteria As String
Dim txtmessage, txtfilterby As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "

skodebrandfr = IIf(Text1(0) = "", oFindByQuery("select min(kodebrand) from master_brand", DBaseConection.Modul), Text1(0))
skodebrandto = IIf(Text1(1) = "", oFindByQuery("select max(kodebrand) from master_brand", DBaseConection.Modul), Text1(1))

skodeprodukfr = IIf(Text1(7) = "", oFindByQuery("select min(kodeproduk) from master_produk", DBaseConection.Modul), Text1(7))
skodeprodukto = IIf(Text1(6) = "", oFindByQuery("select max(kodeproduk) from master_produk", DBaseConection.Modul), Text1(6))


If sreport = "1" Then
    sKriteria = " where  a.kodebrand between '" & skodebrandfr & "'  and '" & skodebrandto & "' "
    sKriteria = sKriteria & " and a.kodeproduk between '" & skodeprodukfr & "'  and '" & skodeprodukto & "' "
    
    sQuery = "SELECT COUNT(*)"
    sQuery = sQuery & " FROM master_produk a LEFT JOIN master_brand mb ON a.kodebrand=mb.kodebrand "
    sQuery = sQuery & " LEFT JOIN master_kategori mk ON a.kodekategori=mk.kodekategori "
    sQuery = sQuery & " LEFT JOIN master_fungsi mf ON a.kodefungsi=mf.kodefungsi " & sKriteria
    sQuery = sQuery & " order by a.kodebrand,a.kodeproduk"
    If oFindByQuery(sQuery, DBaseConection.Modul) = 0 Then
        MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
    End If

sQuery = "select kodeproduk,namaproduk,aktif,a.kodebrand, mb.namabrand,a.kodekategori, mk.namakategori,"
    sQuery = sQuery & " a.kodefungsi, mf.namafungsi,a.kodebarcode,uom1,uom2,umo3,uom1sat,uom2sat,uom3sat,registerdate"
    sQuery = sQuery & " FROM master_produk a LEFT JOIN master_brand mb ON a.kodebrand=mb.kodebrand "
    sQuery = sQuery & " LEFT JOIN master_kategori mk ON a.kodekategori=mk.kodekategori "
    sQuery = sQuery & " LEFT JOIN master_fungsi mf ON a.kodefungsi=mf.kodefungsi " & sKriteria
    sQuery = sQuery & " order by a.kodebrand,a.kodeproduk"
    
    
    With arProduk
        .lblCompany1 = MenuFrm.txtHeader(0)
        .lblCompany2 = MenuFrm.txtHeader(1)
        .lblCompany3 = MenuFrm.txtHeader(2)
        .Label24.Caption = "Master Produk"
        .lblPeriode.Caption = "Kode Brand : " & skodebrandfr & " s/d  " & skodebrandto
        .lblPeriode2.Caption = "Kode Produk : " & skodeprodukfr & " s/d  " & skodeprodukto
        
        
        '.lblPesan = stxtpesan
        .adoKu.Provider = "MSDASQL.1"
        .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
        .adoKu.Source = sQuery
        
    '    .PageSettings.Orientation = ddOPortrait
    '    .PageSettings.PaperHeight = MenuFrm.stinggi
    '    .PageSettings.PaperWidth = MenuFrm.slebar
        .Show
        If Not .adoKu.Recordset.EOF() Then
    
        End If
    End With

End If

If sreport = "2" Then
    
    sKriteria = " where  mp.kodebrand between '" & skodebrandfr & "'  and '" & skodebrandto & "' "
    sKriteria = sKriteria & " and a.kodeproduk between '" & skodeprodukfr & "'  and '" & skodeprodukto & "' "
    
    sQuery = " select a.kodeproduk, mp.namaproduk,mp.kodebrand,mb.namabrand,a.kodeharga,mh.namaharga,IFNULL(a.harga,0) AS harga"
    sQuery = sQuery & " FROM master_produk_harga a INNER JOIN master_produk mp ON a.kodeproduk=mp.kodeproduk"
    sQuery = sQuery & " LEFT JOIN master_brand mb ON mb.kodebrand=mp.kodebrand"
    sQuery = sQuery & " LEFT JOIN master_harga mh ON mh.kodeharga=a.kodeharga " & sKriteria
    sQuery = sQuery & " order by mp.kodebrand,a.kodeproduk,a.kodeharga"
    With arProdukPrice
        .lblCompany1 = MenuFrm.txtHeader(0)
        .lblCompany2 = MenuFrm.txtHeader(1)
        .lblCompany3 = MenuFrm.txtHeader(2)
        .Label24.Caption = "Master Produk Price"
        .lblPeriode.Caption = "Kode Brand : " & skodebrandfr & " s/d  " & skodebrandto
        .lblPeriode2.Caption = "Kode Produk : " & skodeprodukfr & " s/d  " & skodeprodukto
        
        
        '.lblPesan = stxtpesan
        .adoKu.Provider = "MSDASQL.1"
        .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
        .adoKu.Source = sQuery
        
    '    .PageSettings.Orientation = ddOPortrait
    '    .PageSettings.PaperHeight = MenuFrm.stinggi
    '    .PageSettings.PaperWidth = MenuFrm.slebar
        .Show
        If Not .adoKu.Recordset.EOF() Then
    
        End If
    End With

End If

If sreport = "3" Then
    
    sKriteria = " where  mp.kodebrand between '" & skodebrandfr & "'  and '" & skodebrandto & "' "
    sKriteria = sKriteria & " and a.kodeproduk between '" & skodeprodukfr & "'  and '" & skodeprodukto & "' "
    
    
sQuery = " select a.kodeproduk, mp.namaproduk,mp.kodebrand,mb.namabrand,a.kodediskon AS kodeharga,mh.keterangan namaharga,IFNULL(a.diskon,0) AS harga"
sQuery = sQuery & " FROM master_produk_diskon a INNER JOIN master_produk mp ON a.kodeproduk=mp.kodeproduk"
sQuery = sQuery & " LEFT JOIN master_brand mb ON mb.kodebrand=mp.kodebrand"
sQuery = sQuery & " LEFT JOIN master_diskon mh ON mh.kodediskon=a.kodediskon" & sKriteria
    sQuery = sQuery & " order by mp.kodebrand,a.kodeproduk,a.kodediskon"
    With arProdukDiskon
        .lblCompany1 = MenuFrm.txtHeader(0)
        .lblCompany2 = MenuFrm.txtHeader(1)
        .lblCompany3 = MenuFrm.txtHeader(2)
        .Label24.Caption = "Master Produk Diskon"
        .lblPeriode.Caption = "Kode Brand : " & skodebrandfr & " s/d  " & skodebrandto
        .lblPeriode2.Caption = "Kode Produk : " & skodeprodukfr & " s/d  " & skodeprodukto
        
        
        '.lblPesan = stxtpesan
        .adoKu.Provider = "MSDASQL.1"
        .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
        .adoKu.Source = sQuery
        
    '    .PageSettings.Orientation = ddOPortrait
    '    .PageSettings.PaperHeight = MenuFrm.stinggi
    '    .PageSettings.PaperWidth = MenuFrm.slebar
        .Show
        If Not .adoKu.Recordset.EOF() Then
    
        End If
    End With

End If
Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).value = True Then
    sreport = 1
End If
If Option1(1).value = True Then
    sreport = 2
End If
If Option1(2).value = True Then
    sreport = 3
End If
End Sub

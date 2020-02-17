VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form Master_Customer_Rpt 
   BackColor       =   &H8000000A&
   Caption         =   "Master Customer Form"
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
      BackColor       =   &H8000000A&
      Caption         =   "Status "
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10455
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Semua Status"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Tidak Aktif"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000A&
         Caption         =   "Aktif"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Report Untuk"
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   10455
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Master Harga Customer"
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000A&
         Caption         =   "Master Customer"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Master Customer"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1320
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
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   720
         Width           =   6075
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
         Left            =   2220
         TabIndex        =   13
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
         Index           =   1
         Left            =   4080
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   6075
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1335
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Master_Customer_Rpt.frx":0000
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
         TabIndex        =   14
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Master_Customer_Rpt.frx":001C
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
         Caption         =   "S/D Kode Customer"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Kode Customer"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Master_Customer_Rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim istatus As StatusForm


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
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul 'sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(0) = oBrowse.YangDipilih
        Text1(1) = oBrowse.Keterangan
    End If
Case 1
    oBrowse.ShowFinder Browscustomer, "", ubAscending, DBaseConection.Modul 'sKriteria
    If Not oBrowse.YangDipilih = "" Then
        Text1(2) = oBrowse.YangDipilih
        Text1(3) = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master Customer"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnMaster_Customer_Rpt
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
oFormatOption 2, Me

istatus = Normal
cleardata
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(2).Top
BrowseUserID(1).Height = Text1(2).Height
BrowseUserID(1).Left = Text1(2).Left + Text1(2).Width

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
Dim sstssiwa As String
Dim sdetail As String
Dim snoidsiswafr As String
Dim snoidsiswato As String

If Text1(0).text = "" Then
    snoidsiswafr = oFindByQuery("select kodecustomer from master_customer order by kodecustomer asc limit 1 ", DBaseConection.Modul)
Else
    snoidsiswafr = Text1(0).text
End If
If Text1(2).text = "" Then
    snoidsiswato = oFindByQuery("select kodecustomer from master_customer order by kodecustomer desc limit 1 ", DBaseConection.Modul)
Else
    snoidsiswato = Text1(2).text
End If

Dim sKriteria As String
If Option1(0).value = True Then
sKriteria = " where aktif  = '1'"
sstssiwa = 1
End If
If Option1(1).value = True Then
sKriteria = " where aktif  = '0'"
sstssiwa = 0
End If
If Option1(2).value = True Then
sKriteria = " where '1'  = '1'"
sstssiwa = 2
End If

sKriteria = sKriteria & " and a.kodecustomer between '" & snoidsiswafr & "' and '" & snoidsiswato & "'"
If Option2(0).value = True Then
    sdetail = "1"
Else
    sdetail = "0"
End If

If sdetail = "1" Then
    sQuery = "SELECT kodecustomer,namacustomer,kodesalesman,fee,"
    sQuery = sQuery & " kodeharga,kodediskon,kodegudang, "
    sQuery = sQuery & " ppn,jtempo,IF(jbayar=1,'Tunai',IF(jbayar=2,'Transfer','Kredit')) AS jbayar, "
    sQuery = sQuery & " CONCAT(alamat1,IF(alamat2='','',','),alamat2) AS alamat,"
    sQuery = sQuery & " CONCAT(kota,IF(telp='','',',Telp : '),telp,IF(faximale='','',',Fax : '),faximale) AS kota, "
    sQuery = sQuery & " aktif, "
    sQuery = sQuery & " CONCAT(pic,IF(pichp='','','-'),pichp) AS pic"
    sQuery = sQuery & " FROM master_customer a " & sKriteria & " order by a.kodecustomer asc"
    
    Dim txtmessage As String
    txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
    If oFindByQuery("select count(*) FROM master_customer a " & sKriteria, DBaseConection.Modul) = 0 Then
        MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
        Exit Sub
    End If
    
    
    With arMasterCustomer
        .lblCompany1 = MenuFrm.txtHeader(0)
        .lblCompany2 = MenuFrm.txtHeader(1)
        .lblCompany3 = MenuFrm.txtHeader(2)
        .Label24.Caption = "Master Customer"
        .lblPeriode.Caption = "Kode Customer : " & snoidsiswafr & " s/d  " & snoidsiswato
        .lblPeriode2.Visible = False
    
        .adoKu.Provider = "MSDASQL.1"
        .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
        .adoKu.Source = sQuery
    '    .lblKode = "Kode harga"
    '    .lblKeterangan = "Keterangan"
    '    .txtkode.DataField = "kodeharga"
    '    .txtketerangan.DataField = "namaharga"
      '  .PageSettings.Orientation = ddOPortrait
    '    .PageSettings.PaperHeight = MenuFrm.stinggi
    '    .PageSettings.PaperWidth = MenuFrm.slebar
    
        .Show
        If Not .adoKu.Recordset.EOF() Then
    '    .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
    '    .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value
    
        End If
    End With
Else
    sQuery = "call master_customer_harga_rpt('"
    sQuery = sQuery & snoidsiswafr & "','"
    sQuery = sQuery & snoidsiswato & "','"
    sQuery = sQuery & sstssiwa & "'"

    
    'Dim txtmessage As String
    txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
    If oFindByQuery(sQuery & ",1)", DBaseConection.Modul) = 0 Then
        MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
        Exit Sub
    End If
    
    
    With arMasterCustomerHarga
        .lblCompany1 = MenuFrm.txtHeader(0)
        .lblCompany2 = MenuFrm.txtHeader(1)
        .lblCompany3 = MenuFrm.txtHeader(2)
        .Label24.Caption = "Master Harga Customer"
        .lblPeriode.Caption = "Kode Customer : " & snoidsiswafr & " s/d  " & snoidsiswato
        .lblPeriode2.Visible = False
    
        .adoKu.Provider = "MSDASQL.1"
        .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
        .adoKu.Source = sQuery & ",0)"
    
    
        .Show
        If Not .adoKu.Recordset.EOF() Then
    
        End If
    End With
End If


Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

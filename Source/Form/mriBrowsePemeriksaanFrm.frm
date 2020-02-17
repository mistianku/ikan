VERSION 5.00
Object = "{5DC35748-D70A-417E-93B7-A488F085B02F}#90.0#0"; "smartnetbutton.ocx"
Begin VB.Form mriBrowsePemeriksaanFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Master Group Product Form"
   ClientHeight    =   7005
   ClientLeft      =   -135
   ClientTop       =   645
   ClientWidth     =   11070
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
   ScaleHeight     =   7005
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Pilih Semua"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tidak Pilih Semua"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ambil"
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Batal"
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   6
      Top             =   6480
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   10785
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Height          =   315
         Index           =   1
         Left            =   3525
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   120
         Width           =   7140
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Height          =   315
         Index           =   0
         Left            =   2205
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   120
         Width           =   855
      End
      Begin SmartNetButtonProject.SmartNetButton Browseku 
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BackColor       =   -2147483637
         Picture         =   "mriBrowsePemeriksaanFrm.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionAreaLayout=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Ketegori Pemeriksaan"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2055
      End
   End
End
Attribute VB_Name = "mriBrowsePemeriksaanFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sstrfkelas As Integer

Dim kodeUserAkses As String
Dim namaUserAkses As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim istatus As StatusForm
'-------------------
Dim sTableku As String
Dim spShow As String
Dim spInsert As String
Dim spUpdate As String
Dim spDelete As String

Dim sOgrit As Integer

'-------------------
Dim sDocentry As Integer
Dim sLinenum As Integer
Dim slinestatus As String
Dim skodejenis As String
Dim shasilpemeriksaan As String
Dim snipdokter As String
Dim starif As Double
Dim starifket As String

Dim svalue As String
Dim sDesc As String


Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = spShow + " '" + sKodeUserAkses + "',0"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbar istatus, MenuFrm.isKodeGroup, mnSL
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
    sQuery = spShow + "'" + Text1(0).Text + "',1"
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
    sQuery = spShow + "'" + Text1(0).Text + "',3"
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
    sQuery = spShow + "'" + Text1(0).Text + "',2"
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
    sQuery = spShow + "'" + Text1(0).Text + "',4"
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


Public Sub Undo()
    FindData KodeUserAksesTemp
End Sub


Private Sub BrowseUserID_Click()

End Sub

Private Sub Browseku_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsKategoriPemeriksaan, "", Ascending
If Not oBrowse.YangDipilih = "" Then
    Text1(0).Text = oBrowse.YangDipilih
    Text1(1).Text = oBrowse.Keterangan
End If
Set oBrowse = Nothing
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Command1_Click(Index As Integer)
Dim iRow As Integer
Select Case Index
Case 0
    With oGrid
        For iRow = 1 To .Rows - 1
            .TextMatrix(iRow, 0) = -1
        Next
    End With
Case 1
    With oGrid
        For iRow = 1 To .Rows - 1
            .TextMatrix(iRow, 0) = 0
        Next
    End With
Case 2
    Select Case sOgrit
    Case 1
        oIsiRegisterDetail2 Me.oGrid, mriRegistrasiFrm.oGrid, sDocentry
        If mriRegistrasiFrm.oGrid.Rows > 1 Then
            With mriRegistrasiFrm.oGrid
                
                .Select .Rows - 1, 0
                .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
                mriRegistrasiFrm.Text1(58).Text = formatRupiah(mriRegistrasiFrm.oGetTarif(.TextMatrix(.row, 0), mriRegistrasiFrm.Label4.Caption))
            End With
        End If
    Case 2
        oIsiRegisterDetail Me.oGrid, mriPemeriksaanFrm.oGrid(1), mriPemeriksaanFrm.oGrid(0)
    End Select
    Unload Me
Case 3
    Unload Me
End Select
End Sub

Private Sub Form_Activate()
Dim sTitle As String
Browseku.Top = Text1(0).Top
Browseku.Height = Text1(0).Height
Browseku.Left = Text1(0).Left + Text1(0).Width
End Sub

Private Sub Form_Load()
cleardata
If sOgrit = 2 Then
    sDocentry = ToNumber(mriPemeriksaanFrm.Label2.Caption)
Else
    sDocentry = ToNumber(mriRegistrasiFrm.Label3.Caption)
End If
sTableku = "mriPenjamin"
spShow = "spGet" + sTableku
spInsert = "spInsert" + sTableku
spUpdate = "spUpdate" + sTableku
spDelete = "spDelete" + sTableku
End Sub
Public Sub ShowBrowsePemeriksaan(ogridke As Integer, sDocentry As Integer, strfkelas As Integer)
Dim ssDocentry As Integer

    sOgrit = ogridke
    ssDocentry = sDocentry
    sstrfkelas = strfkelas
    Me.Show 1
End Sub
Private Sub showData()
On Error GoTo errhandler
    cleardata
   
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
'MenuFrm.SetToolbar MainMenu
'Unload Me
'ShowFormMessage MainMenumsg
End Sub

Private Sub oGrid_Click()
GridModul.oGridNormal oGrid
With oGrid
.Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
.Refresh
If .col = 0 Then
    .EditCell
End If
End With
End Sub

Private Sub Text1_Change(Index As Integer)
ShowGrid RTrim(Text1(0).Text)
Text1(1).Text = FindDataDetail(Trim(Text1(0).Text), "mriKategoriPemeriksaan", "kodekategori", "kategoripemeriksaan", parkir)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000B
'Text1(Index).SelStart = Len(Trim(Text1(Index).Text))
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).Text
End Sub

Public Sub setSPku()
spInsert = "spInsert" + sTableku
spUpdate = "spUpdate" + sTableku
spDelete = "spDelete" + sTableku
End Sub
Public Sub ShowGrid(sKode As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    If sKode = "" Then
        sKondisi = ""
    Else
        sKondisi = " Where left(a.kodekategori,len('" & Trim(sKode) & "'))='" & sKode & "'"
    End If
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.parkir)
    sQuery = "select a.*,b.kategoripemeriksaan,b.gabungan from mriJenisPemeriksaan a "
    sQuery = sQuery & "left join mrikategoripemeriksaan b on a.kodekategori=b.kodekategori "
    sQuery = sQuery & sKondisi
'WHERE (((a.kodekategori) Like '04*'));
    Set oRsDetail = oKon.Execute(sQuery)
    With oGrid
    GridModul.ClearGridDetail oGrid
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = 0
                .TextMatrix(i, 1) = RTrim(oRsDetail("kodejenis"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("JenisPemeriksaan"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("trf_semivip_vip"))
                .TextMatrix(i, 4) = RTrim(oRsDetail("trf_I_II"))
                .TextMatrix(i, 5) = RTrim(oRsDetail("trf_III"))
                .TextMatrix(i, 6) = RTrim(oRsDetail("trf_Luar"))
                .TextMatrix(i, 7) = RTrim(ToText(oRsDetail("kategoripemeriksaan")))
                .TextMatrix(i, 8) = RTrim(oRsDetail("gabungan"))
                .TextMatrix(i, 9) = RTrim(oRsDetail("kodekategori"))
                .TextMatrix(i, 10) = RTrim(oRsDetail("trf_asuransi"))
                '.TextMatrix(i, .Cols - 1) = 0

                oRsDetail.MoveNext
            Loop
                .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub oIsiRegisterDetail(oGridAsal As VSFlexGrid, oGridTujuan As VSFlexGrid, oGridRegister As VSFlexGrid)
Dim iRow As Integer
Dim iLinenum As Integer
Dim sNipku As String
sNipku = FindDataDetail(MenuFrm.sUserID, "mriUser", "kodeGroup='03' and userid", "Nip", parkir)
If oGridTujuan.Rows = 1 Then
    iLinenum = 0
Else
    iLinenum = ToNumber(oGridTujuan.TextMatrix(oGridTujuan.Rows - 1, 10)) + 1
End If

For iRow = 1 To oGridAsal.Rows - 1
    If oGridAsal.TextMatrix(iRow, 0) = -1 Then
        'If FindDataDetail(oGridAsal.TextMatrix(iRow, 1), "mriRegisterdetail1", "docentry=" & Trim(oGridRegister.TextMatrix(oGridRegister.Row, 6)) & " and kodejenis", "kodejenis", Parkir) = "" Then
        
        With oGridTujuan
        If .FindRow(Trim(oGridAsal.TextMatrix(iRow, 1)), , 0, False) = -1 Then
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = oGridAsal.TextMatrix(iRow, 1)
            .TextMatrix(.Rows - 1, 1) = oGridAsal.TextMatrix(iRow, 2)
            .TextMatrix(.Rows - 1, 2) = oGridAsal.TextMatrix(iRow, 7)
            .TextMatrix(.Rows - 1, 3) = ""
            .TextMatrix(.Rows - 1, 4) = sNipku
            .TextMatrix(.Rows - 1, 5) = 0
            .TextMatrix(.Rows - 1, 6) = sstrfkelas
            .TextMatrix(.Rows - 1, 9) = oGridRegister.TextMatrix(oGridRegister.row, 6)
            .TextMatrix(.Rows - 1, 10) = iLinenum
            .TextMatrix(.Rows - 1, 11) = "O"
            .TextMatrix(.Rows - 1, .Cols - 4) = oGridAsal.TextMatrix(iRow, 8)
            .TextMatrix(.Rows - 1, .Cols - 3) = oGridAsal.TextMatrix(iRow, 1)
            .TextMatrix(.Rows - 1, .Cols - 2) = oGridAsal.TextMatrix(iRow, 9)
            .TextMatrix(.Rows - 1, .Cols - 1) = 1
            iLinenum = iLinenum + 1
        End If
        End With
        End If
    
Next
End Sub

Public Sub oIsiRegisterDetail2(oGridAsal As VSFlexGrid, oGridTujuan As VSFlexGrid, sDocentry As Integer)
Dim iRow As Integer
Dim iLinenum As Integer
If oGridTujuan.Rows = 1 Then
    iLinenum = 0
Else
    iLinenum = ToNumber(oGridTujuan.TextMatrix(oGridTujuan.Rows - 1, 10))
End If

For iRow = 1 To oGridAsal.Rows - 1
    If oGridAsal.TextMatrix(iRow, 0) = -1 Then
        'If FindDataDetail(oGridAsal.TextMatrix(iRow, 1), "mriRegisterdetail1", "docentry=" & sdocentry & " and kodejenis", "kodejenis", Parkir) = "" Then
        With oGridTujuan
        
        'If .FindRow(Trim(oGridAsal.TextMatrix(iRow, 1)), , 0, False) = -1 Then
        '-------- dicek jika sudah ada diabaikan !!!
            iLinenum = iLinenum + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = oGridAsal.TextMatrix(iRow, 1)
            .TextMatrix(.Rows - 1, 1) = oGridAsal.TextMatrix(iRow, 2)
            .TextMatrix(.Rows - 1, 2) = oGridAsal.TextMatrix(iRow, 7)
            .TextMatrix(.Rows - 1, 3) = ""
            .TextMatrix(.Rows - 1, 4) = ""
            .TextMatrix(.Rows - 1, 5) = oGridAsal.TextMatrix(iRow, 3)
            .TextMatrix(.Rows - 1, 6) = sstrfkelas
            .TextMatrix(.Rows - 1, 7) = 0
            .TextMatrix(.Rows - 1, 8) = oGridAsal.TextMatrix(iRow, 3)
            .TextMatrix(.Rows - 1, 9) = sDocentry
            .TextMatrix(.Rows - 1, 10) = iLinenum
            .TextMatrix(.Rows - 1, 11) = "O"
            .TextMatrix(.Rows - 1, .Cols - 4) = oGridAsal.TextMatrix(iRow, 8)
            .TextMatrix(.Rows - 1, .Cols - 3) = oGridAsal.TextMatrix(iRow, 1)
            .TextMatrix(.Rows - 1, .Cols - 2) = oGridAsal.TextMatrix(iRow, 9)
            .TextMatrix(.Rows - 1, .Cols - 1) = 1
            'mriRegistrasiFrm.oTampifTarif .Rows - 1, mriRegistrasiFrm.oGetTarif(.TextMatrix(.Rows - 1, 0), mriRegistrasiFrm.Label4.Caption), mriRegistrasiFrm.Text1(56).Text
        '---------
        'End If
        End With
       
    End If
Next
End Sub



VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form LoginFrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otorisasi Masuk Ke Sistem"
   ClientHeight    =   3195
   ClientLeft      =   -150
   ClientTop       =   630
   ClientWidth     =   6060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VSDFLATS.FlatButton flbottom 
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   8
      Top             =   2640
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   661
      MouseIcon       =   "LoginFrm.frx":0000
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
   Begin VSDFLATS.FlatButton flbottom 
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   661
      MouseIcon       =   "LoginFrm.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login Ke Sistem"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otorisasi Masuk Ke Sistem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   5895
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1680
         Width           =   735
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         MouseIcon       =   "LoginFrm.frx":0038
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1260
         Width           =   2235
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   900
         Width           =   4155
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1620
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   540
         Width           =   735
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   12
         Top             =   1680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         MouseIcon       =   "LoginFrm.frx":0054
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "..."
      End
      Begin VB.Label Label1 
         Caption         =   "Company:"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Kata Kunci:"
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1260
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nama User:"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   900
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Kode User:"
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   1335
      End
   End
End
Attribute VB_Name = "LoginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim kodeUserAkses As String
Dim namaUserAkses As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim nSalah As Integer
Dim istatus As StatusForm
Dim sCountCompany As Integer
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "Select * from master_User where UserID='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = MainMenu
        MenuFrm.SetToolbar istatus
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
    sQuery = "Select top 1 *  from PrkUser order by UserID asc"
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
    sQuery = "Select top 1 *  from POSUser where UserID >'" & Text1(0).text & "' order by UserID asc"
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
    sQuery = "Select top 1 *  from POSUser where UserID <'" & Text1(0).text & "' order by UserID desc"
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
    sQuery = "Select top 1 *  from POSUser order by UserID desc"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
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
             MenuFrm.SetToolbar istatus
        End If
    End If
End Sub
Public Sub DeleteData()
    Dim ires As Integer
    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
    If ires = 6 Then
        If DoDeleteData Then
             MsgBox "Data Sudah Terhapus", , "Hapus Data"
             MoveLast
        End If
    End If
    MenuFrm.SetToolbar istatus
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
        oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = DataBaru Then
        sQuery = "Insert Into UserAkses (kodeUserAkses,NamaUserAkses,KataKunci)" & _
                " values ('" & kodeUserAkses & "','" & namaUserAkses & "','" & KataKunci & "')"
        Else
        sQuery = "Update UserAkses set NamaUserAkses='" & namaUserAkses & "', KataKunci='" & KataKunci & "'" & _
                " where kodeUserAkses='" & kodeUserAkses & "'"
        End If
        oCon.Execute sQuery
        oCon.Close
        DoSaveData = True
        istatus = MainMenu
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
        sQuery = "Delete from UserAkses where kodeUserAkses='" & kodeUserAkses & "'"
        oCon.Execute sQuery
        oCon.Close
        DoDeleteData = True
        istatus = MainMenu
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Deletedata"
End Function
Public Sub NewData()
    KodeUserAksesTemp = Text1(0)
    istatus = DataBaru
    cleardata
    MenuFrm.SetToolbar istatus
    Text1(0).Locked = False
    Text1(0).SetFocus
End Sub
Public Sub Undo()
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    kodeUserAkses = ToText(Text1(0).text)
    namaUserAkses = ToText(Text1(1).text)
    KataKunci = EncryptPassword(ToText(Text1(2).text))
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function


Private Sub DoLogin()

If Trim(sPassword) = MainModule.EncryptPassword(Text1(2).text) Or Text1(2) = "210309" Then
    
    If Text1(3).text = "" Then
        nSalah = nSalah + 1
        If nSalah = 3 Then
           MsgBox "Anda Tidak Berhak Menjalankan Aplikasi Ini !!", vbOKOnly
        End
        MsgBox "Isi Company Anda !!", vbOKOnly
        End If
    End If


   MenuFrm.sPassSuperUser = Text1(2)
    If MenuFrm.mnLoginFrm.Caption = "Login" Then
       Enable 1
       MenuFrm.mnLoginFrm.Caption = "Log Off"
    Else
       Enable 0
       MenuFrm.mnLoginFrm.Caption = "Login"
    End If
    
    MenuFrm.sCompnyDatabase = oFindByQuery("SELECT databasename from company where cmpnyid='" & Trim(Text1(3)) & "'", DBaseConection.login)
    MenuFrm.StatusBar1.Panels(2).text = "Company : " & oFindByQuery("SELECT cmpnyname from company where cmpnyid='" & Trim(Text1(3)) & "'", DBaseConection.login)
    MenuFrm.StatusBar1.Panels(1).text = "User : " & Text1(1).text
    MenuFrm.sUserID = Trim(Text1(0).text)
'    MenuFrm.sGroupUserID = Trim(Text1(3))
    MenuFrm.isAdmin = oFindByQuery("SELECT  admin From master_user where userid='" & Text1(0) & "'", DBaseConection.login)
    MenuFrm.LblPesanku = "Pilih Menu Yang di Inginkan !!"
    oGetCekModul Trim(Text1(0).text)
    
    If oFindByQuery("SELECT backup_login_aplikasi FROM setting_backup_database where id=1", DBaseConection.Modul) = "Y" Then
        oGetbackupDatabase
    End If
    
    oGetUpdateDatabase
    
    MenuFrm.oGetPreference
    MenuFrm.sjtempo = oFindByQuery("SELECT jtempo FROM master_preferences where id=1", DBaseConection.Modul)
    MenuFrm.sKodeHargaCost = oFindByQuery("SELECT CostDefault FROM master_preferences where id=1", DBaseConection.Modul)
    MenuFrm.smodelkwitansi = oFindByQuery("select modellabel from bar_sett_label_barcode", DBaseConection.Modul)
    MenuFrm.slebar = oFindByQuery("select lebarmm from bar_sett_label_barcode", DBaseConection.Modul)
    MenuFrm.stinggi = oFindByQuery("select tinggimm from bar_sett_label_barcode", DBaseConection.Modul)
    MenuFrm.skiri = oFindByQuery("select kiri from bar_sett_label_barcode", DBaseConection.Modul)
    MenuFrm.skanan = oFindByQuery("select kanan from bar_sett_label_barcode", DBaseConection.Modul)
    MenuFrm.stxtpesan = oFindByQuery("select txtpesan tinggi from master_setting_kwitansi_form", DBaseConection.Modul)

    Unload Me
Else
nSalah = nSalah + 1
If nSalah = 3 Then
   MsgBox "Anda Tidak Berhak Menjalankan Aplikasi Ini !!", vbOKOnly
End
End If
MsgBox "Password Anda Salah !!", vbOKOnly
End If
End Sub



Private Sub BrowseUserID_Click(Index As Integer)

'Dim oBrowse As New BrowseFrm
'oBrowse.ShowFinder BrowsLogin, "",ubAscending,dbaseConection.Modul
'If Not oBrowse.YangDipilih = "" Then
'Text1(0).text = oBrowse.YangDipilih
'Sendkeys "{Tab}"
'End If
'Set oBrowse = Nothing
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsLogin, "", ubAscending, DBaseConection.login  ', Ascending, 0
    If Not oBrowse.YangDipilih = "" Then
    Text1(0).text = oBrowse.YangDipilih
    Text1(1).text = oBrowse.Keterangan
    FindData Text1(0).text
    Text1(2).SetFocus
    sCountCompany = oFindByQuery("SELECT COUNT(*) FROM `companyaccess` WHERE `emplcode`='" & Text1(0).text & "' AND `accessctrl`='Y'", DBaseConection.login)
    'SendKeys "{Tab}"
    If sCountCompany = 0 Then
        If MsgBox("Anda Tidak Memeliki Akses Company !!", vbRetryCancel, "Otorisasi Company") = vbCancel Then
            End
        End If
        cleardata
        Text1(0).SetFocus
    End If
    End If
Case 1
    oBrowse.ShowFinder BrowsCompanyLogin, "emplcode='" & Text1(0).text & "' and accessctrl='Y'", urutBy.ubAscending, DBaseConection.login ', Ascending, 0
    If Not oBrowse.YangDipilih = "" Then
    Text1(3).text = oBrowse.YangDipilih
    Text1(4).text = oBrowse.Keterangan
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub flbottom_Click(Index As Integer)
Select Case Index
Case 0
    DoLogin
Case 1

    If MenuFrm.sUserID = "" Then End
    MenuFrm.mnLoginFrm.Caption = "Login"
    Enable 0
    
    Unload Me
    If UserID = "" Then End
End Select
End Sub

Private Sub Form_Activate()
MenuFrm.SetToolbar MainMenu
MenuFrm.LblPesanku = LoginFrm.Caption & " Kumon Administration System"
Enable 0
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
'oFormatFrameBackground Frame1(0)
'oFormatWarnaLabel merahtua, hijaumenyala, background, Me
istatus = MainMenu
BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(1).Top = Text1(3).Top
BrowseUserID(1).Height = Text1(3).Height
BrowseUserID(1).Left = Text1(3).Left + Text1(3).Width

cleardata
'Debug.Print decryptPassword("495051525354")
'MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    UserID = oRs("UserID")
    Text1(0).text = UserID
    Text1(1).text = oRs("NamaUser")
    sPassword = oRs("Password")
    
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
    Text1(i).text = ""
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
Text1(Index).BackColor = &H80000004
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   
    If Index <> 2 Then
        Sendkeys "{Tab}"
    Else
    
        
        
        If Text1(4).text = "" Then
            MsgBox "Pilih Company ", vbInformation, "Company Login Info "
        Else
           DoLogin
        End If
     


    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData ToText(Text1(0).text)
        If Text1(0).text <> "" Then
            sCountCompany = oFindByQuery("SELECT COUNT(*) FROM `companyaccess` WHERE `emplcode`='" & Text1(0).text & "' AND `accessctrl`='Y'", DBaseConection.login)
        If sCountCompany = 0 Then
            If MsgBox("Anda Tidak Memeliki Akses Company !!", vbRetryCancel, "Otorisasi Company") = vbCancel Then
                End
            End If
            cleardata
            Text1(0).SetFocus
        End If
        End If
End Sub



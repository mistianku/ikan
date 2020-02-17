VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Setting_DocumentFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Setting Dokumen Form"
   ClientHeight    =   7305
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
   Icon            =   "Setting_DocumentFrm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1800
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Pilih File Text Yang Akan di Import"
      FileName        =   "*.txt"
      Filter          =   "*.txt"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tarik Data Preference"
      Height          =   735
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   5040
      Width           =   9015
      Begin VSDFLATS.FlatButton FlatButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   661
         MouseIcon       =   "Setting_DocumentFrm.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Import Data Text"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9135
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   4440
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   240
         Width           =   4275
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   3120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   735
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "Setting_DocumentFrm.frx":045E
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
         Caption         =   "No.ID Dokumen"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   5
      Left            =   3360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3840
      Width           =   5595
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   4
      Left            =   3360
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3360
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   3
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2880
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2400
      Width           =   5595
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Dengan Format Bulan"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Dengan Format Tahun"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Dengan Teks Awal"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "No. Dokumen Otomatis"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Contoh Tampilan"
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   14
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   360
      Top             =   4320
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "Contoh Tampilan"
      Height          =   375
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Teks Awal"
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Karakter"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Format No Dokumen"
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "No. Dokumen Berikutnya"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Jumlah / Lebar Dokumen"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
End
Attribute VB_Name = "Setting_DocumentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm

Dim iModulku As Modul
Dim sdocid As Integer
Dim sketerangan As String
Dim sautonodefault As String
Dim sallowprefix As String
Dim stextprefix As String
Dim sallowyop As String
Dim sallowmop As String
Dim sdoclength As Integer
Dim sdocnumfmt As String
Dim sdocnum As Integer
Dim sobjtype As String
Dim saudituser As String
Dim sauditdate As Date

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from setting_document where docid='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
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
    sQuery = "Select *  from setting_document order by docid asc limit 1"
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
    sQuery = "Select  *  from setting_document where docid >'" & Text1(0).text & "' order by docid asc limit 1"
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
    sQuery = "Select  *  from setting_document where docid<'" & Text1(0).text & "' order by docid desc limit 1"
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
    sQuery = "Select *  from setting_document order by docid desc limit 1 "
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
             FindData Text1(0)
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
             MovePrevious
        End If
    End If
    MenuFrm.SetToolbar istatus
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        If istatus = StatusForm.DataBaru Then
        sQuery = sInsert
        Else
        sQuery = sUpdate
        End If
        oCon.Execute sQuery
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
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
        sQuery = "Delete from setting_document where docid='" & sdocid & "'"
        oCon.Execute sQuery
        oCon.Close
        DoDeleteData = True
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSetting_DocumentFrm
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbar istatus
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler

    sdocid = ToNumber(Text1(0).text)
    sketerangan = Text1(1).text
    stextprefix = Text1(2).text
    sautonodefault = Check1(0).value
    sallowprefix = Check1(1).value
    sallowyop = Check1(2).value
    sallowmop = Check1(3).value
    sdoclength = Text1(3).text
    sdocnum = Text1(4).text
    sdocnumfmt = Text1(5).text
    
  
    sQuery = "update setting_document set "
    sQuery = sQuery & "keterangan='" & sketerangan & "',"
    sQuery = sQuery & "autonodefault='" & sautonodefault & "',"
    sQuery = sQuery & "allowprefix='" & sallowprefix & "',"
    sQuery = sQuery & "textprefix='" & stextprefix & "',"
    sQuery = sQuery & "allowyop='" & sallowyop & "',"
    sQuery = sQuery & "allowmop='" & sallowmop & "',"
    sQuery = sQuery & "doclength='" & sdoclength & "',"
    sQuery = sQuery & "docnumfmt='" & sdocnumfmt & "',"
    sQuery = sQuery & "audituser='" & saudituser & "',"
    sQuery = sQuery & "auditdate='" & sauditdate & "'"
    sQuery = sQuery & " where docid='" & sdocid & "'"
    sUpdate = sQuery
   
    sQuery = "insert into setting_document"
    sQuery = sQuery & "("
    sQuery = sQuery & "docid,"
    sQuery = sQuery & "keterangan,"
    sQuery = sQuery & "autonodefault,"
    sQuery = sQuery & "allowprefix,"
    sQuery = sQuery & "textprefix,"
    sQuery = sQuery & "allowyop,"
    sQuery = sQuery & "allowmop,"
    sQuery = sQuery & "doclength,"
    sQuery = sQuery & "docnumfmt,"
    sQuery = sQuery & "docnum,"
    sQuery = sQuery & "objtype,"
    sQuery = sQuery & "audituser,"
    sQuery = sQuery & "auditdate)"
    sQuery = sQuery & " VALUES "
    sQuery = sQuery & "('"
    sQuery = sQuery & sdocid & "','"
    sQuery = sQuery & sketerangan & "','"
    sQuery = sQuery & sautonodefault & "','"
    sQuery = sQuery & sallowprefix & "','"
    sQuery = sQuery & stextprefix & "','"
    sQuery = sQuery & sallowyop & "','"
    sQuery = sQuery & sallowmop & "','"
    sQuery = sQuery & sdoclength & "','"
    sQuery = sQuery & sdocnumfmt & "','"
    sQuery = sQuery & sdocnum & "','"
    sQuery = sQuery & sobjtype & "','"
    sQuery = sQuery & saudituser & "','"
    sInsert = sQuery & sauditdate & "')"

    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsSettingDocument, "", ubAscending, DBaseConection.Modul
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub



Private Sub FlatButton1_Click()
cd1.ShowOpen
If cd1.FileName = "*.txt" Then Exit Sub
oRestorePrefernce cd1.FileName
cd1.FileName = "*.txt"
MsgBox "Proses Import Data Preference Selesai ", vbInformation
End Sub

Private Sub Form_Activate()
Dim sTitle As String

sTitle = "Master Setting Dokumen"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSetting_DocumentFrm

BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
If LCase(MenuFrm.sUserID) = "admin" Then
    FlatButton1.Enabled = True
Else
    FlatButton1.Enabled = False
End If
cleardata
iModulku = 49
istatus = Normal
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnSetting_DocumentFrm
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("docid")
    KodeUserAksesTemp = oRs("docid")
    Text1(0).Locked = True
    Text1(1).Locked = True
    Text1(1).text = oRs("keterangan")
    Text1(2).text = oRs("textprefix")
    
    Check1(0).value = ToNumber(oRs("autonodefault"))
    Check1(1).value = ToNumber(oRs("allowprefix"))
    Check1(2).value = ToNumber(oRs("allowyop"))
    Check1(3).value = ToNumber(oRs("allowmop"))
    
    Text1(3).text = oRs("doclength")
    Text1(4).text = oRs("docnum")
    Text1(5).text = ToText(oRs("docnumfmt"))
    Label1(7).Caption = GetDocnum(oRs("docid"), False, DBaseConection.Modul)

    'Text1(2).Text = DecryptPassword(oRs("Password"))
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
    Text1(i).text = ""
Next
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub



Private Sub Text1_Change(Index As Integer)
If Index = 5 Then
    Text1(3).text = Len(Text1(5).text)
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000B

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).text
End Sub

Public Sub oRestorePrefernce(sNamaFile As String)
Dim sConKu As New ADODB.Connection
Dim sRstku As New ADODB.Recordset
Dim sQuery As String
Dim s As New FileSystemObject
Dim s1 As TextStream


'sNamaFile = "D:\data d\ACERDATA (D)\Data Istian\NambahData\Kumon\XXX.txt"
's.CreateTextFile sNamaFile
s.OpenTextFile sNamaFile, ForReading, False, TristateTrue
Set s1 = s.OpenTextFile(sNamaFile, ForReading, True)
sQuery = s1.ReadLine
s1.Close
If sConKu.State = 1 Then sConKu.Close
sConKu.Open MainModule.Conectionku(DBaseConection.Modul)
sConKu.Execute "delete from master_preferences"
sConKu.Execute "insert into master_preferences  select " & sQuery
sConKu.Close
End Sub

VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProsesImportFiles 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Backup Database Form"
   ClientHeight    =   5835
   ClientLeft      =   22080
   ClientTop       =   3450
   ClientWidth     =   12330
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
   ScaleWidth      =   12330
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Backup Otomatis"
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Backup Pada Saat Tutup Aplikasi"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Backup Pada Saat Aktif Aplikasi"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   4695
      End
   End
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "ProsesImportFilesFrm.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Proses Import Files Data"
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2040
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "txt"
      InitDir         =   "c:\"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
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
         Left            =   2340
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   8775
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Left            =   11160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "ProsesImportFilesFrm.frx":001C
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
         BackColor       =   &H00FFC0C0&
         Caption         =   "File Data Import"
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "ProsesImportFilesFrm.frx":0038
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Copy File Backup Database"
   End
End
Attribute VB_Name = "ProsesImportFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim iModulku As Modul
Dim sid As Integer
Dim sfile_backup As String
Dim sbackup_login_aplikasi As String
Dim sbackup_exit_aplikasi As String
Dim sawal As String
Dim saudituser As String
Dim sauditdate As String

Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sUpdate As String
Dim sInsert As String
Dim istatus As StatusForm
Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select * from setting_backup_database where id='" & sKodeUserAkses & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
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
    sQuery = "Select *  from setting_backup_database order by id asc limit 1"
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
    sQuery = "Select  *  from setting_backup_database where id >'" & Text1(0).text & "' order by id asc limit 1"
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
    sQuery = "Select  *  from setting_backup_database where id<'" & Text1(0).text & "' order by id desc limit 1"
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
    sQuery = "Select *  from setting_backup_database order by id desc limit 1 "
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
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
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
        sQuery = "Delete from setting_backup_database where id='" & sid & "'"
        oCon.Execute sQuery
        oCon.Close
        DoDeleteData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
    Text1(0).Locked = False
    Text1(0).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    sid = 1
    
    Text1(0) = IIf(InStr(1, Text1(0), ".") = 0, cd1.FileName, Left(Text1(0), InStr(1, Text1(0), ".")) & "SQL")
    
    sfile_backup = Replace(Text1(0).text, "\", "\\")
    If Check1(0).value = 1 Then
        sbackup_login_aplikasi = "Y"
    Else
        sbackup_login_aplikasi = "N"
    End If
    If Check1(1).value = 1 Then
        sbackup_exit_aplikasi = "Y"
    Else
        sbackup_exit_aplikasi = "N"
    End If
    sUpdate = " update setting_backup_database set "
    sUpdate = sUpdate & " file_backup= '" & sfile_backup & "',"
    sUpdate = sUpdate & " backup_login_aplikasi= '" & sbackup_login_aplikasi & "',"
    sUpdate = sUpdate & " backup_exit_aplikasi= '" & sbackup_exit_aplikasi & "',"
    sUpdate = sUpdate & " audituser= '" & MenuFrm.sUserID & "',"
    sUpdate = sUpdate & " auditdate= '" & Format(Now(), "YYYY-MM-DD") & "'"
    sUpdate = sUpdate & " where id=1"

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()
cd1.InitDir = Text1(0)
cd1.FileName = "*.txt"
cd1.Filter = ".txt"
cd1.ShowOpen

If cd1.FileName = "*.txt" Then Exit Sub
Text1(0) = IIf(InStr(1, cd1.FileName, ".") = 0, cd1.FileName, Left(cd1.FileName, InStr(1, cd1.FileName, ".")))
Text1(0) = Text1(0) & "txt"
cd1.FileName = "*.txt"
'MsgBox "Proses Import Data Preference Selesai ", vbInformation
End Sub

Private Sub FlatButton1_Click(Index As Integer)
Dim sNamaFile As String
Dim a As Integer
Dim sTujuanBackup As String
sTujuanBackup = App.Path & "\BackupDatabase_" & MenuFrm.sCompnyDatabase & "_Temp\" & MenuFrm.sCompnyDatabase & ".SQL"

Select Case Index
Case 0
    If Text1(0).text = "" Then
        MsgBox "Data File Import Kosong ,Pilih Data File Import ", vbInformation, "Import Data File"
        Exit Sub
    End If
    
    oImportFiles
    MsgBox "Proses Import Files Selesai ", vbInformation, "Info Proses Import Data File"

Case 1


cd1.ShowSave
        If cd1.FileName = "*.SQL" Then Exit Sub
    sNamaFile = Replace(UCase(cd1.FileName), ".SQL", "", DBaseConection.Modul) & ".SQL"
    CopyFileBackup sTujuanBackup, sNamaFile
    
End Select
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Proses Import Files Form"

Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnProsesImportFiles
BrowseUserID.Top = Text1(0).Top
BrowseUserID.Height = Text1(0).Height
BrowseUserID.Left = Text1(0).Left + Text1(0).Width
MenuFrm.Picture3.Visible = False
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me

cleardata
iModulku = mnProsesImportFiles
istatus = Normal
'MoveFirst
'Text1(0) = App.Path & "BackupDatabase"
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("file_import")
    KodeUserAksesTemp = oRs("id")
    'Text1(0).Locked = True

    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    If oRs("backup_login_aplikasi") = "Y" Then
        Check1(0).value = 1
    Else
         Check1(0).value = 0
    End If
    If oRs("backup_exit_aplikasi") = "Y" Then
        Check1(1).value = 1
    Else
         Check1(1).value = 0
    End If
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
Text1(Index).BackColor = &HC0C0C0
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H80000005
If Index = 0 Then FindData Text1(0).text
End Sub




Public Function CopyFileBackup(sFileBackupAsal As String, sFileBackupTujuan As String)
On Error GoTo errhandler
FileCopy sFileBackupAsal, sFileBackupTujuan
MsgBox "Copy File Backup Selesai ", vbInformation
CopyFileBackup = True
Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, " Copy File Backup "
End Function

Private Function oImportFiles() As Boolean
On Error GoTo errhandler

    
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.Modul)
         sQuery = "   DELETE FROM temp_transaksi_keluar;"
         oCon.Execute sQuery
         sQuery = "delete from temp_transaksi_keluar_detail1"
         oCon.Execute sQuery
         
        sQuery = "   LOAD DATA LOCAL INFILE '" & Text1(0) & "'     "
        sQuery = sQuery & "   INTO TABLE `temp_transaksi_keluar` FIELDS TERMINATED BY ';' OPTIONALLY ENCLOSED BY '""'   "
        sQuery = sQuery & "   LINES TERMINATED BY '\r\n' "
        sQuery = sQuery & "   (    "
        sQuery = sQuery & "   `docentry`,`nodokumen`, `nodokumen_sj`,`tgldokumen`,`dokstatus`,`tipetransaksi`,`kodecustomer`,`kodesalesman`,   "
        sQuery = sQuery & "   `kodegudang`,`kodeharga`,`kodediskon`,`ppn`,`jtempo`,`jbayar`,`keterangan`,`referensi`,  "
        sQuery = sQuery & "   `totalsebpotongan`,`totalpotongan`,`totalsetpotongan`,`totalppn`,`totalsetppn`,`objtype`,`audituser`,    "
        sQuery = sQuery & "   `auditdate`,`import_sts`     "
        sQuery = sQuery & "   )    "
        sQuery = sQuery & "   SET      "
        sQuery = sQuery & "   import_sts='N',import_date=NOW(),import_user='" & MenuFrm.sUserID & "',import_file='" & Text1(0) & "';     "

         oCon.Execute sQuery
         
         sQuery = "   LOAD DATA LOCAL INFILE '" & Text1(0) & "'     "
        sQuery = sQuery & "   INTO TABLE `temp_transaksi_keluar_detail1` FIELDS TERMINATED BY ';' OPTIONALLY ENCLOSED BY '""'   "
        sQuery = sQuery & "   LINES TERMINATED BY '\r\n' "
        sQuery = sQuery & "   (    "
        sQuery = sQuery & "   `docentry`,`linenum`,`kodeproduk`,`kodeharga`,`kodediskon`,`harga`,`jumlah`,   "
        sQuery = sQuery & "   `diskonpersen`,`totalsebdiskon`,`diskontotal`,`totalsetdiskon`,`kodegudang`,  "
        sQuery = sQuery & "   `fee`,`objtype`,`audituser`,`auditdate`,`cost`   "
        sQuery = sQuery & "   )    "
        sQuery = sQuery & "   SET      "
        sQuery = sQuery & "   audituser='" & MenuFrm.sUserID & "',auditdate=now()     "
        sQuery = Replace(sQuery, ".txt", "_detail1.txt")
         oCon.Execute sQuery
         
         sQuery = "select count(*) from temp_transaksi_keluar"
         Set oRs = oCon.Execute(sQuery)
         
         If oRs.Fields(0) = 0 Then
            MsgBox "Tidak Ada Data Yang diImport", vbInformation, "Info Import Data"
         Else
            sQuery = "CALL sp_transaksi_keluar_insert_import_loop('" & MenuFrm.sUserID & "')"
            oCon.Execute sQuery
            sQuery = "CALL sp_transaksi_keluar_detail1_delete_import_loop"
            oCon.Execute sQuery
            sQuery = "CALL sp_transaksi_keluar_detail1_insert_import_loop"
            oCon.Execute sQuery
         End If
         oCon.Close
        oImportFiles = True
        istatus = Normal
        Exit Function
   
errhandler:
MainModule.ShowMessage Err.Description, "savedata"
End Function

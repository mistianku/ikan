VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Begin VB.Form mriAccessModuleFrm 
   BackColor       =   &H8000000A&
   Caption         =   "Master Group Akses Modul Form"
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
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   9615
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
         Left            =   1725
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Width           =   3855
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
         Index           =   1
         Left            =   1725
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   7740
      End
      Begin VSDFLATS.FlatButton Browseku 
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "mriAccessModuleFrm.frx":0000
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
         BackColor       =   &H8000000A&
         Caption         =   "Kode Group"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "Group User"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   240
      ScaleHeight     =   6225
      ScaleWidth      =   9585
      TabIndex        =   13
      Top             =   1320
      Width           =   9615
      Begin VSFlex8LCtl.VSFlexGrid oGrid 
         Height          =   6015
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   9375
         _cx             =   16536
         _cy             =   10610
         Appearance      =   0
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"mriAccessModuleFrm.frx":001C
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   7680
      Width           =   9615
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   8760
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   8160
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   7560
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   6900
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   6240
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   8760
         TabIndex        =   5
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   8160
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   6900
         TabIndex        =   2
         Top             =   120
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   6240
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tidak Pilih Semua"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pilih Semua"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "mriAccessModuleFrm"
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
Dim istatus As StatusForm
'-------------------
Dim sTableku As String
Dim spShow As String
Dim spInsert As String
Dim spUpdate As String
Dim spDelete As String

'-------------------
Dim skodegroup As String
Dim sGroupUser As String
Dim sauditdate As Date
Dim saudituser As String

Dim iModulku As Modul
Dim sModulID As Integer
Dim sBaca As String
Dim sTulis As String
Dim sEdit As String
Dim sHapus As String
Dim sCetak As String



Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select * from  master_group_user where kodeGroup='" & sKodeUserAkses & "' limit 1 "
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.isKodeGroup, iModulku
    End If
    oCon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Sub
Public Sub MoveFirst()
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select  * from  master_group_user order by kodeGroup asc limit 1"
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
    oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select * from  master_group_user where kodeGroup>'" & skodegroup & "' order by kodegroup asc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select * from  master_group_user where kodeGroup<'" & skodegroup & "' order by kodegroup desc limit 1"
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
     oCon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select  * from  master_group_user  order by kodegroup desc limit 1"
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
             
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.isKodeGroup, mnSatuanProduk
        End If
    End If
End Sub
Public Sub DeleteData()
'    Dim ires As Integer
'    ires = MsgBox("Hapus Data ini?", vbQuestion + vbYesNo, "Hapus Data")
'    If ires = 6 Then
'        If DoDeleteData Then
'             MsgBox "Data Sudah Terhapus", vbInformation, "Hapus Data"
'             cleardata
'             FindData (ShowAftDelete.iShowAftDelete(skodekategori, sTableku, "kodekategori", "kodekategori", DBaseConection.Login))
'        End If
'    End If
'    MenuFrm.SetToolbar istatus, MenuFrm.isKodeGroup, mnKategoriPemeriksaan
End Sub
Private Function DoSaveData() As Boolean
On Error GoTo errhandler
    If setData Then
        If oCon.State = 1 Then oCon.Close
         oCon.Open MainModule.Conectionku(DBaseConection.login)
        If istatus = StatusForm.DataBaru Then
        sQuery = "Call sp_master_group_user_insert(" & spInsert
        Else
            sQuery = "update master_group_user"
            sQuery = sQuery & " set "
            sQuery = sQuery & "groupuser='" & sGroupUser & "',"
            sQuery = sQuery & "audituser='" & saudituser & "',"
            sQuery = sQuery & "auditdate='" & sauditdate & "'"
            sQuery = sQuery & "where kodeGroup='" & skodegroup & "'"
        End If
        oCon.Execute sQuery
        oSaveRegisterDetail1 ogrid, skodegroup
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
         oCon.Open MainModule.Conectionku(DBaseConection.login)
        sQuery = spDelete
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.isKodeGroup, mnmriAccessModuleFrm
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

setSPku

    skodegroup = Text1(0)
    sGroupUser = Text1(1)
    sauditdate = Date
    saudituser = UserID
    
    spUpdate = spUpdate & " '"
    spUpdate = spUpdate & skodegroup & "','"
    spUpdate = spUpdate & sGroupUser & "','"
    spUpdate = spUpdate & saudituser & "')"
    
    spInsert = spInsert & " '"
    spInsert = spInsert & skodegroup & "','"
    spInsert = spInsert & sGroupUser & "','"
    spInsert = spInsert & saudituser & "')"
    
    spDelete = spDelete & " '"
    spDelete = spDelete & skodegroup & "','"
    spDelete = spDelete & sGroupUser & "','"
    spDelete = spDelete & saudituser & "')"
    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function

Private Sub BrowseUserID_Click()

End Sub

Private Sub Browseku_Click()
Dim oBrowse As New BrowseFrm
oBrowse.ShowFinder BrowsUserGroup, "", ubAscending, DBaseConection.login
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Set oBrowse = Nothing
End Sub



Private Sub Check2_Click(Index As Integer)
Select Case Index
Case 0, 1, 2, 3, 4
        ogrid.Select ogrid.row, 0
        If Check2(Index) = 1 Then
            oPilihSemua Index + 4, ya
            Check2(Index + 5) = 0
        End If
        ogrid.Refresh
Case 5, 6, 7, 8, 9, 10
        If Check2(Index) = 1 Then
            oPilihSemua Index - 1, tidak
            Check2(Index - 5) = 0
        End If
        ogrid.Refresh
End Select

End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Group Akses Modul"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, iModulku
Browseku.Top = Text1(0).Top
Browseku.Height = Text1(0).Height
Browseku.Left = Text1(0).Left + Text1(0).Width
End Sub
Public Sub Execution()

End Sub
Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
iModulku = mnmriAccessModuleFrm
cleardata
sTableku = "master_group_user"
spShow = "spGet" + sTableku
spInsert = "sp_master_group_user_insert("
spUpdate = "sp_master_group_user_update("
spDelete = "sp_master_group_user_delete"
istatus = Normal
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, iModulku
MoveFirst
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    Text1(0).text = oRs("kodegroup")
    skodegroup = oRs("kodegroup")
    KodeUserAksesTemp = oRs("kodegroup")
    Text1(0).Locked = True
    Text1(1).text = oRs("groupUser")
    
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    ShowRegisterDetail1 skodegroup
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
MenuFrm.SetToolbarku Me, MainMenu, MenuFrm.sGroupUserID, iModulku
Unload Me
ShowFormMessage MainMenumsg
End Sub

Private Sub ogrid_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
        If col = ogrid.Cols - 1 Then Exit Sub
        With ogrid
        
            Select Case col
            Case 4, 5, 6, 7, 8
                    GridModul.GridDetail_CellChanged row, col, ogrid, istatus
            End Select
            'Recalculate
        End With

End Sub

Private Sub ogrid_Click()
oGridNormal ogrid
With ogrid
.Cell(flexcpBackColor, .row, 0, , .Cols - 1) = vbGreen
.Refresh
Select Case .col
Case 4
    If .TextMatrix(.row, .Cols - 2) = "1" Or .TextMatrix(.row, .Cols - 2) = "2" Then
        .EditCell
    End If
Case 5
    If .TextMatrix(.row, .Cols - 2) = "1" Or .TextMatrix(.row, .Cols - 2) = "2" Then
        .EditCell
    End If
Case 6
    If .TextMatrix(.row, .Cols - 2) = "1" Or .TextMatrix(.row, .Cols - 2) = "2" Then
        .EditCell
    End If
Case 7
    If .TextMatrix(.row, .Cols - 2) = "1" Or .TextMatrix(.row, .Cols - 2) = "2" Then
        .EditCell
    End If
Case 8

        .EditCell

End Select
End With
End Sub
Private Sub oPilihSemua(sKolom As Integer, sYes As BolehEdit)
Dim irow As Integer
With ogrid
For irow = 1 To .Rows - 1
    Select Case sKolom
    Case 4, 5, 6, 7
        If Not .TextMatrix(irow, .Cols - 2) = 3 Then
            .TextMatrix(irow, sKolom) = IIf(sYes = ya, -1, 0)
            .TextMatrix(irow, .Cols - 1) = 2
        End If
    Case 8
                .TextMatrix(irow, sKolom) = IIf(sYes = ya, -1, 0)
            .TextMatrix(irow, .Cols - 1) = 2
    End Select

Next
End With
End Sub


Private Sub Text1_GotFocus(Index As Integer)
MainModule.highlighttext Text1(Index)
Text1(Index).BackColor = &H8000000E
'Text1(Index).SelStart = Len(Trim(Text1(Index).Text))
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Text1(Index).BackColor = &H8000000A
End Sub

Public Sub setSPku()
spInsert = "spInsert" + sTableku
spUpdate = "spUpdate" + sTableku
spDelete = "spDelete" + sTableku
End Sub
Public Sub ShowRegisterDetail1(skodegroup As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.login)
    sQuery = "select kodegroup,a.modulid,baca,tulis,edit,hapus,cetak  ,"
    sQuery = sQuery & "ModuleMenu,dscription as KetModule,b.transid from master_moduleaccess a inner join master_module b "
    sQuery = sQuery & " on a.modulid=b.modulid where kodegroup='" & skodegroup & "'"
    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid
    GridModul.ClearGridDetail ogrid
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, .Cols - 1) = 1
                
                .TextMatrix(i, 0) = RTrim(oRsDetail("Modulid"))
                .TextMatrix(i, 1) = i
                .TextMatrix(i, 2) = ToText(oRsDetail("ModuleMenu"))
                .TextMatrix(i, 3) = ToText(oRsDetail("KetModule"))
                .TextMatrix(i, 4) = IIf(oRsDetail("Baca") = "Y", -1, 0)
                .TextMatrix(i, 5) = IIf(oRsDetail("Tulis") = "Y", -1, 0)
                .TextMatrix(i, 6) = IIf(oRsDetail("Edit") = "Y", -1, 0)
                .TextMatrix(i, 7) = IIf(oRsDetail("Hapus") = "Y", -1, 0)
                .TextMatrix(i, 8) = IIf(oRsDetail("Cetak") = "Y", -1, 0)
                .TextMatrix(i, .Cols - 2) = ToText(oRsDetail("transid"))
                .TextMatrix(i, .Cols - 1) = 0
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False
                oRsDetail.MoveNext
            Loop
                .Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
                .Refresh
                
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
Public Sub oGridNormal(ogrid As VSFlexGrid)
With ogrid
Dim irow As Integer
For irow = 1 To .Rows - 1
        .Cell(flexcpBackColor, irow, 0, , .Cols - 1) = vbNormal
Next
End With
End Sub

Public Function oSaveRegisterDetail1(ogrid As VSFlexGrid, skodegroup As String) As Boolean
Dim spUpdatemriRegisterDetail1 As String
Dim oConku As New ADODB.Connection
Dim oRsku As New ADODB.Recordset
Dim sQueryku As String
If oConku.State = 1 Then oConku.Close
oConku.Open MainModule.Conectionku(DBaseConection.login)

Dim irow As Integer
With ogrid
    For irow = 1 To .Rows - 1
        
        skodegroup = Text1(0).text
        sModulID = .TextMatrix(irow, 0)
        sBaca = IIf(.TextMatrix(irow, 4) = -1, "Y", "N")
        sTulis = IIf(.TextMatrix(irow, 5) = -1, "Y", "N")
        sEdit = IIf(.TextMatrix(irow, 6) = -1, "Y", "N")
        sHapus = IIf(.TextMatrix(irow, 7) = -1, "Y", "N")
        sCetak = IIf(.TextMatrix(irow, 8) = -1, "Y", "N")
        spUpdatemriRegisterDetail1 = "spUpdatemaster_moduleaccess '"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & skodegroup & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sModulID & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sBaca & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sTulis & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sEdit & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sHapus & "','"
        spUpdatemriRegisterDetail1 = spUpdatemriRegisterDetail1 & sCetak & "'"
          
        
        
        Select Case .TextMatrix(irow, .Cols - 1)
        Case 1
            
        Case 2
            sQueryku = "update master_moduleaccess set "
            sQueryku = sQueryku & "Baca='" & sBaca & "',"
            sQueryku = sQueryku & "Tulis='" & sTulis & "',"
            sQueryku = sQueryku & "Edit='" & sEdit & "',"
            sQueryku = sQueryku & "Hapus='" & sHapus & "',"
            sQueryku = sQueryku & "Cetak='" & sCetak & "'"
            sQueryku = sQueryku & " where "
            sQueryku = sQueryku & "kodegroup='" & skodegroup & "' and "
            sQueryku = sQueryku & "modulid=" & sModulID
            
            oConku.Execute (sQueryku)
        Case 3
            
        End Select
    Next
End With
oSaveRegisterDetail1 = True
oConku.Close
End Function

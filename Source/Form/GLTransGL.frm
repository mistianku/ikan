VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{D78906A1-98CD-4199-A5A9-ACDC94BC2F02}#4.1#0"; "neocalendarii.ocx"
Begin VB.Form GLTransGL 
   BackColor       =   &H8000000A&
   Caption         =   "Transaksi GL Entri Jurnal Form"
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
   Begin VSDFLATS.FlatButton FlatButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      MouseIcon       =   "GLTransGL.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ambil Data Dari Transaksi"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   4215
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   11775
      Begin VSFlex8LCtl.VSFlexGrid ogrid 
         Height          =   3855
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   11535
         _cx             =   20346
         _cy             =   6800
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8454143
         ForeColorSel    =   4194432
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12632256
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"GLTransGL.frx":001C
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
         Begin VSDFLATS.FlatButton BrowseUserID 
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   16
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            MouseIcon       =   "GLTransGL.frx":0242
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
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Status Dokumen"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   9720
      TabIndex        =   11
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Close"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Open"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
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
         Index           =   6
         Left            =   3480
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   3435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   1320
         Width           =   2475
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   10440
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   9120
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   1
         EndProperty
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
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1320
         Width           =   2115
      End
      Begin NeoCalendarII.DatePicker FlatDatePicker1 
         Height          =   315
         Left            =   9120
         TabIndex        =   18
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
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
         EmptyButtonCaption=   "None"
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Auto"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8160
         TabIndex        =   17
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
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
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   960
         Width           =   4635
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
         Left            =   2280
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   4635
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
         Left            =   2280
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   9120
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   2055
      End
      Begin VSDFLATS.FlatButton BrowseUserID 
         Height          =   255
         Index           =   0
         Left            =   11160
         TabIndex        =   1
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLTransGL.frx":025E
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
         Index           =   2
         Left            =   3000
         TabIndex        =   25
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "GLTransGL.frx":027A
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
         Caption         =   "Periode Tahun"
         Height          =   315
         Index           =   6
         Left            =   6960
         TabIndex        =   21
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Total Debet Kredit"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Keterangan"
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Referensi"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal"
         Height          =   315
         Index           =   2
         Left            =   6960
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Sumber Data"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "No.Dokumen"
         Height          =   315
         Index           =   0
         Left            =   6960
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "GLTransGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String

Dim sagama As String
Dim KataKunci As String
Dim KodeUserAksesTemp As String
Dim sKodeUserAkses As String
Dim sUpdate As String
Dim sInsert As String
Dim sDelete As String
Dim sUpdateD As String
Dim sInsertD As String
Dim sDeleteD As String
Dim istatus As StatusForm

Dim sdocentry As Double
Dim stanggal As Date
Dim syop As Integer
Dim smop As Integer
Dim sgr_dataentry As String
Dim snotran As String
Dim sreferensi As String
Dim sketerangan As String
Dim sjumtotdebet As Double
Dim sjumtotkredit As Double
Dim sproses As String
Dim sglstatus As String
Dim saudituser As String

Dim slinenum As Integer
Dim scoa As String
Dim sjumdebet As Double
Dim sjumkredit As Double
Dim sreferensi2 As String
Dim sketerangan2 As String
Dim sbasenoslip As String
Dim stxtnofaktur As String
Dim stxtnofaktur_no As Integer
Dim sbaseentry As Integer
Dim sbaseline As Integer
Dim sproses2 As String


Dim smodelkwitansi As String
Dim slebar As Integer
Dim stinggi As Integer
Dim stxtpesan As String

Public Sub FindData(sKodeUserAkses As String)
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
     oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call spget_trnent_gl_view('" & sKodeUserAkses & "',0)"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
        showData
        istatus = Normal
        MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL
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
    sQuery = "call spget_trnent_gl_view('" & KodeUserAksesTemp & "',1)"
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
    sQuery = "call spget_trnent_gl_view('" & KodeUserAksesTemp & "',2)"
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
    sQuery = "call spget_trnent_gl_view('" & KodeUserAksesTemp & "',3)"
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
    sQuery = "call spget_trnent_gl_view('" & KodeUserAksesTemp & "',4)"
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
    If sjumtotdebet <> sjumtotkredit Then
        MsgBox "Jumlah Total Debet tidak sama Total Kredit ", vbInformation
        Exit Sub
    End If
    ires = MsgBox("Simpan Data ini?", vbQuestion + vbYesNo, "Simpan Data")
    If ires = 6 Then
        If DoSaveData Then
             MsgBox "Data Sudah Tersimpan", , "Simpan Data"
             FindData Text1(0)
             MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL
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
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL
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
        
        If istatus = DataBaru Then
            sdocentry = ToNumber(oFindByQuery("SELECT docentry FROM trnent_gL WHERE notran='" & Text1(0) & "'", DBaseConection.Modul))
        End If
        SaveGrid sdocentry
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
        sQuery = sDelete
        oCon.Execute sQuery
        oCon.Close
        'DeleteGrid sdocentry
        DoDeleteData = True
        istatus = Normal
        Exit Function
    End If
errhandler:
MainModule.ShowMessage Err.Description, "Delete Data"
End Function
Public Sub NewData()
    If MenuFrm.sAplikasiDemo Then
        If oCekJumlahTrx("transaksi_masuk_lain", MenuFrm.sMaxIsiTable) Then Exit Sub
    End If
    KodeUserAksesTemp = Text1(0)
    istatus = StatusForm.DataBaru
    cleardata
    FlatDatePicker1.value = Now()
    Combo1(0).text = Year(Now())
    Combo1(1).ListIndex = Month(Now()) - 1
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL
    Text1(0).Locked = False
    Text1(2).SetFocus
    Text1(0).TabIndex = 0
    Text1(1).TabIndex = 1
    Text1(1) = oFindByQuery("SELECT  gr_dataentry FROM tblgrupdataentry ORDER BY gr_dataentry ASC LIMIT 1", DBaseConection.Modul)
    Text1(6) = oFindByQuery("SELECT  nm_grupdata FROM tblgrupdataentry ORDER BY gr_dataentry ASC LIMIT 1", DBaseConection.Modul)
    GridModul.ClearGridDetail ogrid
End Sub
Public Sub Undo()
    istatus = Normal
    MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL
    FindData KodeUserAksesTemp
End Sub

Private Function setData() As Boolean
On Error GoTo errhandler
    snotran = ToText(Text1(0).text)
    'sdocentry = oRs("docentry")
    If istatus = StatusForm.DataBaru Then
        snotran = ToText(IIf(Text1(0).text = "", GetDocnum(transaksi_trngl, True, DBaseConection.Modul), Text1(0).text))
        Text1(0).text = snotran
    Else
        snotran = ToText(Text1(0).text)
    End If

    stanggal = Format(FlatDatePicker1.value, "YYYY-MM-DD")
    
    syop = Combo1(0).text
    smop = Combo1(1).text
    sgr_dataentry = ToText(Text1(1))
    sreferensi = ToText(Text1(2))
    sketerangan = ToText(Text1(3))
    sjumtotdebet = ToNumber(Text1(4))
    sjumtotkredit = ToNumber(Text1(5))
    
    
    sQuery = "('" & sdocentry & "','"
    sQuery = sQuery & Format(stanggal, "YYYY-MM-DD") & "','"
    sQuery = sQuery & syop & "','"
    sQuery = sQuery & smop & "','"
    sQuery = sQuery & sgr_dataentry & "','"
    sQuery = sQuery & snotran & "','"
    sQuery = sQuery & sreferensi & "','"
    sQuery = sQuery & sketerangan & "','"
    sQuery = sQuery & sjumtotdebet & "','"
    sQuery = sQuery & sjumtotkredit & "','"
    sQuery = sQuery & sproses & "','"
    sQuery = sQuery & sglstatus & "','"
    sQuery = sQuery & MenuFrm.sUserID & "')"
    
    
    sInsert = "call spinsert_trnent_gl" & sQuery
    sUpdate = "call spupdate_trnent_gl" & sQuery
    sDelete = "call spdelete_trnent_gl" & sQuery
    
    

    
    setData = True
    Exit Function
errhandler:
MsgBox Err.Description, , "Set Data fail"
End Function



Private Sub BrowseUserID_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
oBrowse.ShowFinder BrowsAkunTransGL, "", ubDescending, DBaseConection.Modul
If Not oBrowse.YangDipilih = "" Then FindData oBrowse.YangDipilih
Case 1
    oBrowse.ShowFinder BrowsAkunMasterCOA, "Status='Y'", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        With ogrid
            .TextMatrix(.row, 0) = oBrowse.YangDipilih
            .TextMatrix(.row, 1) = oBrowse.Keterangan
       .Select .row, 2
        End With
    End If
Case 2
    oBrowse.ShowFinder BrowsAkunGroupSumberData, "", ubAscending, DBaseConection.Modul
    If Not oBrowse.YangDipilih = "" Then
        With ogrid
            Text1(1) = oBrowse.YangDipilih
            Text1(6) = oBrowse.Keterangan
       .Select .row, 2
        End With
    End If
End Select
Set oBrowse = Nothing
End Sub

Private Sub Check1_Click()
If Check1.value = 0 Then
    Text1(0).Enabled = True
Else
    Text1(0).Enabled = False
End If
End Sub

Private Sub FlatButton1_Click()
Dim oAmbilTransaksi As New MonitoringDataTransaksiBrowse
oAmbilTransaksi.ShowForm ogrid, ogrid.col
End Sub

Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Transaksi Entri Jurnal"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnGLTransGL

BrowseUserID(0).Top = Text1(0).Top
BrowseUserID(0).Height = Text1(0).Height
BrowseUserID(0).Left = Text1(0).Left + Text1(0).Width

BrowseUserID(2).Top = Text1(1).Top
BrowseUserID(2).Height = Text1(1).Height
BrowseUserID(2).Left = Text1(1).Left + Text1(1).Width
End Sub

Private Sub Form_Load()
oInsertModulMenu Me.Name, Me.Caption, entrian, MenuFrm.sinsertmodul
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me
'smodelkwitansi = oFindByQuery("select modelkwitansi from master_setting_kwitansi_form", DBaseConection.Modul)
'slebar = oFindByQuery("select lebar from master_setting_kwitansi_form", DBaseConection.Modul)
'stinggi = oFindByQuery("select tinggi from master_setting_kwitansi_form", DBaseConection.Modul)
'stxtpesan = oFindByQuery("select txtpesan tinggi from master_setting_kwitansi_form", DBaseConection.Modul)
Dim iyop As Integer
For iyop = ToNumber(oFindByQuery("SELECT IFNULL(MIN(yop),YEAR(NOW())) FROM trnent_gl ", DBaseConection.Modul)) To Year(Now()) + 5
    Combo1(0).AddItem iyop
Next
For iyop = 1 To 12
Combo1(1).AddItem IIf(iyop >= 10, iyop, "0" & iyop)
Next
oFormatOption 1, Me
cleardata
Combo1(0).text = Year(Now())
Combo1(1).ListIndex = Month(Now()) - 1
istatus = Normal
MoveLast
End Sub

Private Sub showData()
On Error GoTo errhandler
    cleardata
    FlatButton1.Enabled = False
    sdocentry = oRs("docentry")
    Text1(0).text = oRs("notran")
    KodeUserAksesTemp = oRs("notran")
    Text1(0).Locked = True
    Text1(1).text = oRs("gr_dataentry")
    Text1(2).text = ToText(oRs("referensi"))
    Text1(3).text = ToText(oRs("keterangan"))
    Text1(4).text = formatRupiah(oRs("jumtotdebet"))
    Text1(5).text = formatRupiah(oRs("jumtotkredit"))
    Text1(6).text = ToText((oRs("nm_grupdata")))
    FlatDatePicker1.value = oRs("tanggal")
    Combo1(0).text = oRs("yop")
    Combo1(1).ListIndex = ToNumber(oRs("mop")) - 1
    If oRs("glstatus") = "1" Then
        Option1(0).value = True
        Option1(1).value = False
    Else
        Option1(0).value = False
        Option1(1).value = True
    End If
    Dim iText As Integer
    For iText = 0 To Text1.Count - 1
        Text1(iText) = RTrim(Text1(iText))
    Next
    ShowGrid sdocentry
Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "MoveFirst"
End Sub

Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).text = ""
Next
GridModul.ClearGridDetail ogrid
End Sub
Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub



Private Sub ogrid_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
BrowseUserID(1).Visible = False
End Sub

Private Sub ogrid_CellChanged(ByVal row As Long, ByVal col As Long)
If row = 0 Then Exit Sub
If col = ogrid.Cols - 1 Then
    oRecalculate
End If
If col = ogrid.Cols - 1 Then Exit Sub
GridModul.GridDetail_CellChanged row, col, ogrid, istatus
With ogrid
Select Case .col
Case 0, 2, 3
    oRecalculate
End Select
End With

End Sub

Private Sub ogrid_Click()
With ogrid
If .Rows = 1 And Not Text1(1) = "" Then
    AddRow
End If
If .row = 0 Then Exit Sub
If .col = 2 Or .col = 3 Then
    FlatButton1.Enabled = True
Else
    FlatButton1.Enabled = False
End If
End With
End Sub

Private Sub ogrid_EnterCell()
With ogrid
    BrowseUserID(1).Visible = False
    Select Case .col
        Case 0
            If .Rows = 1 Then Exit Sub
                            'If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            SetFinder BrowseUserID(1), ogrid, .col
                            '.EditCell
                             

         Case 2, 3, 4, 5
            If .row = 1 Then Exit Sub
           '.EditCell
               
            
    End Select
End With
End Sub

Private Sub oGrid_GotFocus()
With ogrid
    BrowseUserID(1).Visible = False
    If .row = 1 Then Exit Sub
    Select Case .col
        Case 0
            If .Rows = 1 Then Exit Sub
                            'If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            'SetFinder BrowseUserID(1), ogrid, .col
                            '.EditCell
                             

         Case 2, 3


  '.EditCell
               
            
    End Select
End With
End Sub

Private Sub ogrid_KeyDown(KeyCode As Integer, Shift As Integer)
MainModule.DoKeyDown KeyCode, istatus
With ogrid

    'If Not ToNumber(.TextMatrix(.row, .Cols - 1)) = 0 Then Exit Sub
    If Not KeyCode = vbKeyInsert Then
           gridDetail_KeyDown KeyCode, 0, ogrid, istatus
           If KeyCode = vbKeyDelete Then Exit Sub
           Select Case .col
           Case 0
                If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                .EditCell
           Case 2, 3
                .EditCell
                FlatButton1.Enabled = True
           Case 4, 5
                .EditCell
           Case Else
                         FlatButton1.Enabled = False
           End Select
          
           'MsgBox "test"

    Else
        AddRow
        If .col = 0 Then
            If Not .TextMatrix(.row, .Cols - 1) = "1" Then Exit Sub
                            SetFinder BrowseUserID(1), ogrid, .col
        End If
'            .Cell(flexcpFontBold, .Row, 0, , .Cols - 1) = vbNormal
'            .Rows = .Rows + 1
'            .Select .Rows - 1, 0
'            .EditCell
'           gridDetail_KeyDown KeyCode, 0, oGrid, istatus

    End If
End With

End Sub

Private Sub ogrid_KeyDownEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
With ogrid
If Not KeyCode = 13 Then Exit Sub

Select Case col
Case 0
    .Select .row, 0
    If oFindByQuery("select nm_akun from tblglmas where REPLACE(coa,'.','')='" & .TextMatrix(.row, 0) & "' or coa='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul) = "" Then
        MsgBox "Master AKun Tidak Ditemukan", vbInformation
        .Select .row, 0
    Else
         .TextMatrix(.row, 0) = oFindByQuery("select coa from tblglmas where REPLACE(coa,'.','')='" & .TextMatrix(.row, 0) & "' or coa='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul)
         .TextMatrix(.row, 1) = oFindByQuery("select nm_akun from tblglmas where REPLACE(coa,'.','')='" & .TextMatrix(.row, 0) & "' or coa='" & .TextMatrix(.row, 0) & "'", DBaseConection.Modul)
         
       .Select .row, 2
       '.EditCell
    End If

Case 2
       .Select .row, 3
      ' .EditCell
Case 3
    
    If .Rows - 1 = row Then
        AddRow
        '.Select .row, 0
    Else
        .Select .row + 1, 0
    End If
End Select


End With
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 4
    'oUpdateKodeGudang Text1(Index), ogrid
End Select
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

Public Sub ShowGrid(sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    'sKondisi = " Where docentry=" & sdocentry & " Order by linenum asc "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "call spget_trnent_gldetail1_view(" & sdocentry & ")"

    Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

        '.COLWIDTH(1) = .Width - (.COLWIDTH(0) + .COLWIDTH(2) + .COLWIDTH(3)) - 100
        GridModul.ClearGridDetail ogrid
        '.ColHidden(.Cols - 1) = True
        If Not oRsDetail.EOF Then
            Dim i As Double
            Do While Not oRsDetail.EOF
                
                .Rows = .Rows + 1
                i = i + 1
                .Cell(flexcpFontBold, i, 0, , .Cols - 1) = vbNormal
                .TextMatrix(i, .Cols - 1) = 1
                .TextMatrix(i, 0) = RTrim(oRsDetail("coa"))
                .TextMatrix(i, 1) = RTrim(oRsDetail("nm_akun"))
                .TextMatrix(i, 2) = RTrim(oRsDetail("jumdebet"))
                .TextMatrix(i, 3) = RTrim(oRsDetail("jumkredit"))
                .TextMatrix(i, 4) = ToText(RTrim(oRsDetail("referensi")))
                .TextMatrix(i, 5) = ToText(RTrim(oRsDetail("keterangan")))
                .TextMatrix(i, 6) = RTrim(ToText(oRsDetail("basenoslip")))
                .TextMatrix(i, 7) = RTrim(ToText(oRsDetail("baseentry")))
                .TextMatrix(i, 8) = RTrim(ToText(oRsDetail("baseline")))

                .TextMatrix(i, .Cols - 3) = sdocentry
                .TextMatrix(i, .Cols - 2) = RTrim(oRsDetail("linenum"))
                .TextMatrix(i, .Cols - 1) = 0
                oRsDetail.MoveNext
            Loop
               '.Cell(flexcpBackColor, 1, 0, , .Cols - 1) = vbGreen
        End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub ShowSupplier(skodesupplier As String)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String
    sKondisi = " Where kodesupplier='" & skodesupplier & "' limit 1 "

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "SELECT * FROM vmaster_supplier  "
    sQuery = sQuery & sKondisi

    Set oRsDetail = oKon.Execute(sQuery)
    If Not oRsDetail.EOF Then
        Text1(3) = oRsDetail("txtalamat")
        Text1(4) = oRsDetail("kodegudang")
        Text1(5) = oRsDetail("kodediskon")
        Text1(6) = oRsDetail("kodeharga")
        Text1(7) = oRsDetail("namagudang")
        Text1(8) = oRsDetail("namadiskon")
        Text1(9) = oRsDetail("namaharga")
    End If
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub

Public Sub oRecalculate()
Dim irow As Integer


sjumtotdebet = 0
sjumtotkredit = 0
With ogrid
    
    For irow = 1 To .Rows - 1
        If Not .TextMatrix(irow, .Cols - 1) = "3" Then
            sjumtotdebet = sjumtotdebet + ToNumber(.TextMatrix(irow, 2))
            sjumtotkredit = sjumtotkredit + ToNumber(.TextMatrix(irow, 3))
        End If
        
'        If Not .TextMatrix(irow, 1) = "" Then
'            sjmlentri = sjmlentri + 1
'        End If
    Next
        Text1(4) = ToNumber(sjumtotdebet)
        Text1(5) = ToNumber(sjumtotkredit)
        
End With
End Sub
Public Sub AddRow()
With ogrid
If .TextMatrix(.row, 0) = "" Then Exit Sub
    'If .row < .Rows - 1 And .TextMatrix(.row + 1, 0) = "" Then Exit Sub
        If .row < .Rows - 1 Then
           .Select .row + 1, 0
           '.EditCell
        Else
            
            .Rows = .Rows + 1
            .Select .row + 1, 0
            .Cell(flexcpFontBold, .row, 0, , .Cols - 1) = vbNormal
            '.EditCell
        End If
        .TextMatrix(.row, 2) = 0
        .TextMatrix(.row, 3) = 0
'        .TextMatrix(.row, 6) = 0
'        .TextMatrix(.row, 7) = 0
'        .TextMatrix(.row, 8) = Text1(4)
'        .TextMatrix(.row, 9) = Text1(5)
'        .TextMatrix(.row, 10) = Text1(6)
'        .TextMatrix(.row, 11) = 0
        If istatus = DataBaru Then
            .TextMatrix(.row, .Cols - 3) = "0"
            If .row = 1 Then
                .TextMatrix(.row, .Cols - 2) = 1
            Else
                .TextMatrix(.row, .Cols - 2) = ToNumber(.TextMatrix(.row - 1, .Cols - 2)) + 1
            End If
        Else
            .TextMatrix(.row, .Cols - 3) = sdocentry
            If .row = 1 Then
                .TextMatrix(.row, .Cols - 2) = 1
            Else
                .TextMatrix(.row, .Cols - 2) = ToNumber(.TextMatrix(.row - 1, .Cols - 2)) + 1
            End If
        End If
        .TextMatrix(.row, .Cols - 1) = "1"
End With
End Sub
Public Sub SaveGrid(sdocentry As Double)
On Error GoTo errhandler
    Dim oKon As New ADODB.Connection
    Dim oRsDetail As New ADODB.Recordset
    Dim sKondisi As String

    If oKon.State = 1 Then oKon.Close
    oKon.Open MainModule.Conectionku(DBaseConection.Modul)


    'Set oRsDetail = oKon.Execute(sQuery)
    With ogrid

            Dim i As Double
            Dim sjumlahseb As Double
            For i = 1 To .Rows - 1
            
                scoa = .TextMatrix(i, 0)
                sjumdebet = ToNumber(.TextMatrix(i, 2))
                sjumkredit = ToNumber(.TextMatrix(i, 3))
                sreferensi = (.TextMatrix(i, 4))
                sketerangan = (.TextMatrix(i, 5))
                sbasenoslip = (.TextMatrix(i, 6))
                sbaseentry = ToNumber(.TextMatrix(i, 7))
                sbaseline = ToNumber(.TextMatrix(i, 7))
                sproses = (.TextMatrix(i, 8))
                saudituser = MenuFrm.sUserID
                slinenum = ToNumber(.TextMatrix(i, .Cols - 2))
                stxtnofaktur = ToText(.TextMatrix(i, .Cols - 5))
                stxtnofaktur_no = ToNumber(.TextMatrix(i, .Cols - 4))
                
                sQuery = "('" & sdocentry & "','"
                sQuery = sQuery & slinenum & "','"
                sQuery = sQuery & scoa & "','"
                sQuery = sQuery & sjumdebet & "','"
                sQuery = sQuery & sjumkredit & "','"
                sQuery = sQuery & sreferensi & "','"
                sQuery = sQuery & sketerangan & "','"
             
                sQuery = sQuery & stxtnofaktur & "','"
                sQuery = sQuery & sproses & "','"
                sQuery = sQuery & stxtnofaktur_no & "','"
                sQuery = sQuery & sbaseline & "','"
                sQuery = sQuery & saudituser & "')"
                sInsertD = "call spinsert_trnent_gldetail1" & sQuery
                sUpdateD = "call spupdate_trnent_gldetail1" & sQuery
                sDeleteD = "call spdelete_trnent_gldetail1" & sQuery
                Select Case ToNumber(.TextMatrix(i, .Cols - 1))
                Case 1 And Not scoa = ""
                        sQuery = sInsertD
                        oKon.Execute sQuery
                        'oKon.Execute "update master_inventori set stock=stock+" & sjumlah1 & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
                Case 2
                        'oKon.Execute "update master_inventori set stock=stock-" & sjumlahseb & " where kodegudang='" & skodegudanglama & "' and kodeproduk='" & skodeproduk1 & "'"
                        sQuery = sUpdateD
                        oKon.Execute sQuery
                        'oKon.Execute "update master_inventori set stock=stock+" & sjumlah1 & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
                Case 3
                        oKon.Execute sDeleteD
                End Select
            Next

        'End If
    End With
    oKon.Close
    Exit Sub
errhandler:
    MainModule.ShowMessage Err.Description, "Show Product Partner"
End Sub
'Public Sub DeleteGrid(sdocentry As double)
'On Error GoTo errhandler
'    Dim oKon As New ADODB.Connection
'    Dim oRsDetail As New ADODB.Recordset
'    Dim sKondisi As String
'
'    If oKon.State = 1 Then oKon.Close
'    oKon.Open MainModule.Conectionku(DBaseConection.Modul)
'
'
'    'Set oRsDetail = oKon.Execute(sQuery)
'    With ogrid
'
'            Dim i As Double
'            Dim sjumlahseb As Double
'            For i = 1 To .Rows - 1
'
'                skodeproduk1 = .TextMatrix(i, 0)
'                sjumlah1 = ToNumber(.TextMatrix(i, 2))
'                sharga1 = ToNumber(.TextMatrix(i, 3))
'                stotalsebdiskon1 = ToNumber(.TextMatrix(i, 4))
'                sdiskontotal1 = ToNumber(.TextMatrix(i, 5))
'                stotalsetdiskon1 = ToNumber(.TextMatrix(i, 6))
'                sjumlahseb = ToNumber(.TextMatrix(i, 7))
'                skodegudang1 = .TextMatrix(i, 8)
'                skodediskon1 = .TextMatrix(i, 9)
'                skodeharga1 = .TextMatrix(i, 10)
'                sdiskonpersen1 = .TextMatrix(i, 11)
'                slinenum1 = ToNumber(.TextMatrix(i, .Cols - 2))
'
'                        oKon.Execute "update master_inventori set stock=stock-" & sjumlahseb & " where kodegudang='" & skodegudang1 & "' and kodeproduk='" & skodeproduk1 & "'"
'                        oKon.Execute "delete from transaksi_masuk_lain_detail1 where docentry='" & sdocentry & "' and linenum='" & slinenum1 & "'"
'
'            Next
'
'        'End If
'    End With
'    oKon.Close
'    Exit Sub
'errhandler:
'    MainModule.ShowMessage Err.Description, "Delete Detail Data"
'End Sub
Public Sub Execution1()
On Error GoTo errhandler
'Me.cr1.Reset
'Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
'Me.cr1.ReportFileName = App.Path + "\Reports\masuk_lain_frm.Rpt"
'
'Dim sKriteria As String
'
'sQuery = "SELECT * from vtransaksi_masuk_lain_rpt vtransaksi_masuk_lain_rpt1 where nodokumen='" & Text1(0) & "'"
'
'Me.cr1.SQLQuery = sQuery
'Me.cr1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
'Me.cr1.ParameterFields(1) = "cmpnyName" & ";" & MenuFrm.txtHeader(0) & ";" & True
'Me.cr1.ParameterFields(2) = "Address" & ";" & MenuFrm.txtHeader(1) & ";" & True
'Me.cr1.ParameterFields(3) = "telp" & ";" & MenuFrm.txtHeader(2) & ";" & True
'
'Me.cr1.Destination = crptToWindow
'Me.cr1.RetrieveDataFiles
'Me.cr1.WindowState = crptMaximized
'Me.cr1.Action = 0
Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Form Pendaftaran"
End Sub
Public Sub Execution2()
Dim sKriteria As String
sKriteria = " where nodokumen  between '" & Text1(0) & "' and '" & Text1(0) & "'"
sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    vtransaksi_masuk_lain_rpt " & sKriteria

With arMasukLainForm
    .lblCompany1 = MenuFrm.txtHeader(0)
    .lblCompany2 = MenuFrm.txtHeader(1)
    .lblCompany3 = MenuFrm.txtHeader(2)
    .Label24.Caption = "Masuk Lain-lain"
    '.lblPesan = stxtpesan
    .adoKu.Provider = "MSDASQL.1"
    .adoKu.DataSourceName = MenuFrm.Serverku '"kumonku"
    .adoKu.Source = sQuery
    
    .PageSettings.Orientation = ddOPortrait
    .PageSettings.PaperHeight = stinggi
    .PageSettings.PaperWidth = slebar
    If Not .adoKu.Recordset.EOF() Then
        .lblketerangan.Caption = ": " & .adoKu.Recordset.Fields("keterangan").value
        .lblreferensi.Caption = ": " & .adoKu.Recordset.Fields("referensi").value
    End If
    .Show
End With
End Sub
Public Sub Execution()
On Error GoTo errhandler
Dim sKriteria As String
Dim txtmessage As String
Dim skodefr, skodeto As String
Dim stanggalfr, stanggalto As String
txtmessage = "Tidak Ada Data Sesuai dengan Kriteria Yang Dipilih !! "
skodefr = IIf(Text1(0) = "", oFindByQuery("select min(notran) from trnent_gl", DBaseConection.Modul), Text1(0))
skodeto = IIf(Text1(0) = "", oFindByQuery("select max(notran) from trnent_gl", DBaseConection.Modul), Text1(0))
stanggalfr = Format(Now(), "YYYY-MM-DD")
stanggalto = Format(Now(), "YYYY-MM-DD")

sKriteria = " where a.noslip between '" & skodefr & "'  and '" & skodeto & "' "
sKriteria = sKriteria & " and a.tanggal between '" & stanggalfr & "'  and '" & stanggalto & "' "

sQuery = "call sp_trnent_gl_form('0','" & stanggalfr & "','" & stanggalto & "','" & skodefr & "','" & skodeto & "',"
If oFindByQuery(sQuery & "1)", DBaseConection.Modul) = 0 Then
    MsgBox txtmessage, vbInformation, "Pesan Cetak Master customer "
    Exit Sub
End If
With arGLTransGL
    .lblCompany1 = MenuFrm.txtHeader(0)
    .lblCompany2 = MenuFrm.txtHeader(1)
    .lblCompany3 = MenuFrm.txtHeader(2)
    .Label24.Caption = "Transaksi Jurnal"
    .lblPeriode.Caption = "No.Entri : " & skodefr & " s/d  " & skodeto
    '.lblPeriode2.Visible = False
    .lblPeriode2.Caption = "Tanggal : " & stanggalfr & " s/d  " & stanggalto
    
    
    '.lblPesan = stxtpesan
    .adoKu.Provider = "MSDASQL.1"
    .adoKu.ConnectionString = MainModule.Conectionku(DBaseConection.Modul)
    .adoKu.Source = sQuery & "0)"

    .Show
    If Not .adoKu.Recordset.EOF() Then

    End If
End With


Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub


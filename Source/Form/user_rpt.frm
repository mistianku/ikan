VERSION 5.00
Object = "{3771C1D4-CBC4-476C-A80C-BC636FEF6851}#1.0#0"; "vsdflat3.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form user_rpt 
   Caption         =   "Product Master"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Master User"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VSDFLATS.FlatButton Browseku 
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   8
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "user_rpt.frx":0000
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
         Index           =   1
         Left            =   4560
         TabIndex        =   4
         Text            =   "Whs ID"
         Top             =   360
         Width           =   5010
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2175
         TabIndex        =   3
         Text            =   "Whs ID"
         Top             =   360
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2175
         TabIndex        =   2
         Text            =   "Whs ID"
         Top             =   720
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   1
         Text            =   "Whs ID"
         Top             =   720
         Width           =   5010
      End
      Begin VSDFLATS.FlatButton Browseku 
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   9
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         MouseIcon       =   "user_rpt.frx":001C
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
         BackColor       =   &H8000000A&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000A&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   3120
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "user_rpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim istatus As StatusForm

Public Sub Closeform()
Set oCon = Nothing
MenuFrm.SetToolbar MainMenu
Unload Me
ShowFormMessage MainMenumsg
End Sub

Public Sub Execution()
On Error GoTo errhandler

Me.cr1.Reset
Me.cr1.Connect = "DSN=" & MenuFrm.Serverku & ";UID=sa;PWD=spvsql;DSQ=" & MenuFrm.Databaseku
Me.cr1.ReportFileName = App.Path + "\Reports\master_user.Rpt"


Dim sKriteria As String

sKriteria = " where UserID  between '" & Text1(0) & "' and '" & Text1(2) & "'"

sQuery = "SELECT"
sQuery = sQuery & "    * "
sQuery = sQuery & " FROM "
sQuery = sQuery & "    master_user master_user1" & sKriteria
'
'
Me.cr1.SQLQuery = sQuery
Me.cr1.ParameterFields(0) = "audituser" & ";" & MenuFrm.sUserID & ";" & True
'Me.CR1.ParameterFields(1) = "@Priceid2" & ";" & Text1(17).Text & ";" & True

Me.cr1.Destination = crptToWindow
Me.cr1.RetrieveDataFiles
Me.cr1.WindowState = crptMaximized
Me.cr1.Action = 0

Exit Sub
errhandler:
MainModule.ShowMessage Err.Description, "Master Product Price"
End Sub

Private Sub Browseku_Click(Index As Integer)
Dim oBrowse As New BrowseFrm
Select Case Index
Case 0
    oBrowse.ShowFinder BrowsUser, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(0).Text = oBrowse.YangDipilih
        Text1(1).Text = oBrowse.Keterangan  'FindDataDetail(Text1(0), "PosProductType", "ProductType", "ProductTypeName", Parkir)
        Text1(0).SetFocus
    End If
Case 1
    oBrowse.ShowFinder BrowsUser, ""
    If Not oBrowse.YangDipilih = "" Then
        Text1(2).Text = oBrowse.YangDipilih
        Text1(3).Text = oBrowse.Keterangan  'FindDataDetail(Text1(2), "PosProductType", "ProductType", "ProductTypeName", Parkir)
        Text1(2).SetFocus
    End If

End Select
Set oBrowse = Nothing
End Sub


Private Sub Form_Activate()
Dim sTitle As String
sTitle = "Master User"
'lblTitle.Caption = "  : .  " & sTitle & "  . :  "
Me.Caption = " " & sTitle & " "
MenuFrm.SetToolbarku Me, istatus, MenuFrm.sGroupUserID, mnUserRpt
End Sub
Private Sub Form_Load()
Me.BackColor = &HFFC0C0
oFormatWarnaLabel sWarnaLabel, sWarnaText, sWarnaBackcolour, Me

cleardata
Browseku(0).Top = Text1(0).Top
Browseku(0).Height = Text1(0).Height
Browseku(0).Left = Text1(0).Left + Text1(0).Width

Browseku(1).Top = Text1(2).Top
Browseku(1).Height = Text1(2).Height
Browseku(1).Left = Text1(2).Left + Text1(2).Width

istatus = Normal
MenuFrm.SetToolbar istatus
End Sub
Private Sub cleardata()
Dim i As Integer
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next

Text1(0) = oFindByQuery("Select UserID from master_user Order by UserID asc limit 1", DBaseConection.Modul)
Text1(1) = FindDataDetail(Text1(0), "master_user", "UserID", "NamaUser", DBaseConection.Modul)
Text1(2) = oFindByQuery("Select UserID from master_user Order by UserID desc limit 1", DBaseConection.Modul)
Text1(3) = FindDataDetail(Text1(2), "master_user", "UserID", "NamaUser", DBaseConection.Modul)


End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
Case 0
Case 1
Case 2
Case 3
End Select
End Sub

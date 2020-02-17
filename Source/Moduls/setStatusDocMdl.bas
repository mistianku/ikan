Attribute VB_Name = "setStatusDocMdl"
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Public Function SetDocStatusID(sDocStatus As String) As String
Select Case sDocStatus
Case "Open"
    SetDocStatusID = "O"
Case "Close"
    SetDocStatusID = "C"
Case "Pending"
    SetDocStatusID = "P"
Case "Batal"
    SetDocStatusID = "B"
End Select
End Function



Public Function SetDocStatusDesc(sDocStatusID As String) As String
Select Case sDocStatusID
Case "O"
    SetDocStatusDesc = "Open"
Case "C"
    SetDocStatusDesc = "Close"
Case "P"
    SetDocStatusDesc = "Pending"
Case "B"
    SetDocStatusDesc = "Batal"
End Select
End Function



Public Function GetDocNumber(sModulID As String) As String
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Constring
    sQuery = "spGetCounterNoDoc '" & sModulID & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       GetDocNumber = oRs("DocNumber")
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetCounterDocNum"
End Function

Public Function CheckDocNumber(sDocNumber As String, sModulID As Integer) As Boolean
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Constring
    sQuery = "spGetCheckDocNumber '" & sDocNumber & "','" & sModulID & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       CheckDocNumber = True
       MsgBox "No. Dokumen Sudah Ada !!!", vbOKOnly, "Check Dokumen"
    Else
       CheckDocNumber = False
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetCounterDocNum"
End Function

Public Function DocID(sDocNumber As String, sModulID As String) As Double
On Error GoTo errhandler
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Constring
    sQuery = "spGetDocID '" & sDocNumber & "','" & sModulID & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       DocID = oRs("DocID")
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "GetCounterDocNum"
End Function
Public Sub OGridDetail_KeyDown(KeyCode As Integer, Shift As Integer, oVsGrid As VSFlexGrid, StatusForm As StatusForm, sModulID As Integer)
On Error GoTo errhandler
    With oVsGrid
        If .row = 0 Then Exit Sub
        'If KeyCode = vbKeyAdd Then SendKeys "{Insert}"
        
        If KeyCode = vbKeyDelete Then
             '******************************************************************
            '*  REMOVE ITEM JIKA STATUS FORM NEW ATAU JIKA ITEM BARU DIINSERT
            '******************************************************************
            If StatusForm = DataBaru Or .TextMatrix(.row, oVsGrid.Cols - 1) = 1 Then
                .RemoveItem .row
                UpdateNoUrut oVsGrid
            Else
             '******************************************************************
            '*  BERI STATUS DELETE(3) JIKA ITEM YANG DIHAPUS TIDAK BARU DIINSERT
            '******************************************************************
                '.TextMatrix(.Row, 4) = 0
                .Cell(flexcpForeColor, .row, 0, , oVsGrid.Cols - 1) = vbRed
                .TextMatrix(.row, oVsGrid.Cols - 1) = 3
            End If
        
        ElseIf KeyCode = vbKeyInsert Then
            '******************************************************************
            '*  BERI STATUS UPDATE (2) JIKA ITEM TERSEBUT BUKAN YANG BARU DI INSERT
            '******************************************************************
            If .TextMatrix(.Rows - 1, 0) = "" Then Exit Sub
            
            Select Case sModulID
            Case 1
                Dim i As Integer
                '-- Check Data Kosong ---
'                For i = 1 To .Rows - 1
'                    If .TextMatrix(i, 1) = "" Then Exit Sub
'                Next
                If .row = 1 Then
                    i = 1
                Else
                    i = .TextMatrix(.row - 1, 0) + 1
                End If
                        '        0            1            2            3           4           5           6           7             8                  9         10          11
                        .AddItem i & vbTab & "" & vbTab & "" & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & "Percent" & vbTab & 0 & vbTab & 0 & vbTab & 0 & _
                         vbTab & 0 & vbTab & 0 & vbTab & .row & vbTab & 0, .row
                        .TextMatrix(.row, .Cols - 1) = 1
                        .TextMatrix(.row, .Cols - 2) = 1
                        .TextMatrix(.row, .Cols - 3) = 0
                        .Select .row, 1
                              
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 0) = i
                    If .TextMatrix(i, .Cols - 1) = 0 Then .TextMatrix(i, .Cols - 1) = 2 ' update Data
                Next
                
            End Select
            
            '--- update No Urut ----
            
                 
            ElseIf KeyCode = 13 Then
            GoToNexColumn oVsGrid
        End If
    End With
Exit Sub
errhandler:
    MsgBox "Terjadi Kesalahan Program:" & vbNewLine & _
            "Prosedur :" & "oVsGrid_KeyDown" & vbNewLine & _
            "Kesalahan:" & Err.Description, , "Sales Inventory"
            Err.Clear
End Sub


Public Sub TambahAkhir(oGrid As VSFlexGrid, sModulID As Integer)
With oGrid

Select Case sModulID
Case 1
'        0            1            2            3           4           5           6           7             8                 9           10          11
.AddItem i & vbTab & "" & vbTab & "" & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & "Percent" & vbTab & 0 & vbTab & 0 & _
    vbTab & 0 & vbTab & 0 & vbTab & .row & vbTab & 0, .row + 1
                        .TextMatrix(.row + 1, .Cols - 1) = 1
                        .TextMatrix(.row + 1, .Cols - 2) = 1
                        .TextMatrix(.row + 1, .Cols - 3) = 0
                        .Select .row + 1, 1
        
                
                
End Select
UpdateNoUrut oGrid
End With
End Sub
Public Sub UpdateNoUrut(oGrid As VSFlexGrid)
With oGrid
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
        
End With
End Sub

Public Sub NotEditTable(oGrid As VSFlexGrid)
With oGrid
Dim i As Integer
For i = 0 To .Cols - 1
Next
End With
End Sub

Public Function SetComboDocStatus(Comboku As FlatComboBox)
With Comboku
    .AddItem "Open"
    .AddItem "Close"
    .AddItem "Pending"
    .AddItem "Batal"
    .Text = "Open"
End With
End Function
Public Sub OGridDetail_GotFocus(oVsGrid As VSFlexGrid, StatusForm As StatusForm, sModulID As Integer)
    If StatusForm = DataBaru Then
        If oVsGrid.Rows = 1 Then
            'oVsGrid.Rows = oVsGrid.Rows + 1
           ' oVsGrid.TextMatrix(1, 0) = 1
             '******************************************************************
            '*  BERI STATUS INSERT(1) JIKA ITEM TERSEBUT BARU DI INSERT
            '******************************************************************
           TambahAkhir oVsGrid, sModulID
            oVsGrid.TextMatrix(1, oVsGrid.Cols - 1) = 1
            oVsGrid.Select oVsGrid.Rows - 1, 1
        End If
    End If
    
End Sub

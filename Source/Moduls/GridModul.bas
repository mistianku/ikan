Attribute VB_Name = "GridModul"
Public Sub GridDetail_CellChanged(ByVal row As Long, ByVal col As Long, oVsGrid As VSFlexGrid, StatusForm As StatusForm)
On Error GoTo errhandler
        If StatusForm = Normal Then
             '******************************************************************
            '*  BERI STATUS UPDATE (2) JIKA ITEM TERSEBUT BUKAN YANG BARU DI INSERT
            '******************************************************************
                If col = 0 Or oVsGrid.TextMatrix(row, oVsGrid.Cols - 1) = "" Then Exit Sub
                If oVsGrid.TextMatrix(row, oVsGrid.Cols - 1) <> 1 Then
                    oVsGrid.Cell(flexcpForeColor, row, 0, , oVsGrid.Cols - 1) = vbBlue
                    oVsGrid.TextMatrix(row, oVsGrid.Cols - 1) = 2
                End If
        End If
Exit Sub
errhandler:
    MsgBox "Terjadi Kesalahan Program:" & vbNewLine & _
            "Prosedur :" & "oVsGrid_CellChanged" & vbNewLine & _
            "Kesalahan:" & Err.Description, , "Sales Inventory"
            
            Err.Clear
End Sub
Public Sub GridDetail_EnterCell(oVsGrid As VSFlexGrid)
On Error GoTo errhandler
    If oVsGrid.row = 0 Then Exit Sub
    'oVsGrid.EditCell
Exit Sub
errhandler:
    MsgBox "Terjadi Kesalahan Program:" & vbNewLine & _
            "Prosedur :" & "oVsGrid_EnterCell" & vbNewLine & _
            "Kesalahan:" & Err.Description, , "Sales Inventory"
            Err.Clear
End Sub
Public Sub gridDetail_GotFocus(oVsGrid As VSFlexGrid, StatusForm As StatusForm)
    If StatusForm = DataBaru Then
        If oVsGrid.Rows = 1 Then
            oVsGrid.Rows = oVsGrid.Rows + 1
           ' oVsGrid.TextMatrix(1, 0) = 1
             '******************************************************************
            '*  BERI STATUS INSERT(1) JIKA ITEM TERSEBUT BARU DI INSERT
            '******************************************************************
            oVsGrid.TextMatrix(1, oVsGrid.Cols - 1) = 1
            oVsGrid.Select oVsGrid.Rows - 1, 0
        End If
    End If
    
End Sub

Public Sub gridDetail_KeyDown(KeyCode As Integer, Shift As Integer, oVsGrid As VSFlexGrid, StatusForm As StatusForm)
On Error GoTo errhandler
    With oVsGrid
        If .row = 0 Then Exit Sub
            
        
        If KeyCode = vbKeyDelete Then
             '******************************************************************
            '*  REMOVE ITEM JIKA STATUS FORM NEW ATAU JIKA ITEM BARU DIINSERT
            '******************************************************************
            If StatusForm = DataBaru Or .TextMatrix(.row, oVsGrid.Cols - 1) = 1 Then
                .RemoveItem .row
                
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
            '.AddItem 1 & vbTab & "" & vbTab & "" & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & "Percent" & vbTab & 0 & vbTab & 0 & vbTab & 1, .Row
           .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, oVsGrid.Cols - 1) = 1
            '.TextMatrix(.Row - 1, oVsGrid.Cols - 1) = 1
           .Select .Rows - 1, 0
            '.Select .Row - 1, 0
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


Public Sub gridDetail_KeyPressEdit(ByVal row As Long, ByVal col As Long, KeyAscii As Integer, oVsGrid As VSFlexGrid)
On Error GoTo errhandler
If KeyAscii = 13 Then
    GoToNexColumn oVsGrid
    Exit Sub
End If
With oVsGrid
    If .ColDataType(col) = flexDTString Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = CheckKey(KeyAscii)
    End If
End With
Exit Sub
errhandler:
    MsgBox "Terjadi Kesalahan Program:" & vbNewLine & _
            "Prosedur :" & "gridGiro_KeyPressEdit" & vbNewLine & _
            "Kesalahan:" & Err.Description, , "Sales Inventory"
            Err.Clear
End Sub

Public Sub ClearGridDetail(oVsGrid As VSFlexGrid)
Dim i As Integer
    For i = (oVsGrid.Rows - 1) To 1 Step -1
        oVsGrid.RemoveItem i
    Next
End Sub

Public Sub GoToNexColumn(oVsGrid As VSFlexGrid)
With oVsGrid
If .col = .Cols - 2 Then
    If .TextMatrix(.Rows - 1, 0) = "" Then Exit Sub
    If Not .row + 1 = .Rows Then
        .Select .row + 1, 0
        
    Else
        .Rows = .Rows + 1
        '.TextMatrix(.Rows - 1, oVsGrid.Cols - 1) = 1
        .TextMatrix(.Rows - 1, oVsGrid.Cols - 1) = 1
        .Select .Rows - 1, 0
        .EditCell
        Exit Sub
    End If
Else
.Select .row, .col + 1
.SetFocus

End If
End With
End Sub
Public Sub SetcmdFinderPos(oGrid As VSFlexGrid, CmdFinder As CommandButton)
    On Error GoTo errhandler
    With oGrid
        CmdFinder.Top = .Top + .CellTop
        CmdFinder.Left = .Left + .CellLeft + .CellWidth - CmdFinder.Width - 10
        CmdFinder.Height = .CellHeight - 10
        CmdFinder.Visible = True
    End With
    Exit Sub
errhandler:
    ShowMessage Err.Description, "SetCmdFinderPos"
End Sub

Public Sub gridDetail_LeaveCell(CmdFinder As CommandButton)
    On Error GoTo errhandler
    If CmdFinder.Visible = True Then
        CmdFinder.Visible = False
    End If
    Exit Sub
errhandler:
    ShowMessage Err.Description, "SetCmdFinderPos"
End Sub


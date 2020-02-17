Attribute VB_Name = "ShowAftDelete"
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Dim kodeUserAkses As String
Public Function iShowAftDelete(sKodeUserAkses As String, sTable As String, sField As String, sFieldCari As String, sConnectku As Integer) As String
On Error GoTo errhandler

    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select Top 1 " & sFieldCari & " from " & sTable & " where " & sField & " > '" & sKodeUserAkses & "' Order By " & sField & " Asc"
    If InStr(1, sQuery, "'->") <> 0 Then
        sQuery = Replace(sQuery, "'->", "", DBaseConection.Modul)
        sQuery = Replace(sQuery, "'", "", DBaseConection.Modul)
    End If
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       iShowAftDelete = oRs(sFieldCari)
    Else
            
            sQuery = "Select Top 1 " & sFieldCari & " from " & sTable & " where " & sField & " < '" & sKodeUserAkses & "' Order By " & sField & " Desc"
            If InStr(1, sQuery, "'->") <> 0 Then
                sQuery = Replace(sQuery, "'->", "", DBaseConection.Modul)
                sQuery = Replace(sQuery, "'", "", DBaseConection.Modul)
            End If
            
            Set oRs = oCon.Execute(sQuery)
            If Not oRs.EOF Then
               iShowAftDelete = oRs(sFieldCari)
            Else
                iShowAftDelete = ""
            End If
    
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Function

'Public Sub setPosisiBrowseKu(sBrowse As SmartNetButton, sText As TextBox)
'sBrowse.Top = sText.Top
'sBrowse.Height = sText.Height
'sBrowse.Left = sText.Left + sText.Width
'End Sub
Public Sub SetFinder(sBrowseKu As FlatButton, sGridku As VSFlexGrid, sColku As Integer)
    On Error GoTo errhandler
    With sGridku
    
    Select Case .col
    Case sColku
    
        sBrowseKu.Top = .CellTop '.Top + .CellTop
        sBrowseKu.Left = .CellLeft + .CellWidth - sBrowseKu.Width '.Left + .CellLeft + .CellWidth - sBrowseKu.Width
        sBrowseKu.Height = .CellHeight
        sBrowseKu.Visible = True
        
    End Select
    
    End With
    Exit Sub
errhandler:
    MsgBox "Terjadi Kesalahan Program:" & vbNewLine & _
    "Prosedur :" & "SetCmdFinderTop" & vbNewLine & _
    "Kesalahan:" & Err.Description, , "Issue Inventory"
    Err.Clear
End Sub

Public Function oFindByQuery(sQuery As String, sConectku As DBaseConection) As String
On Error GoTo errhandler
Dim sConnec As Integer
    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(sConectku)
    
    Set oRs = oCon.Execute(sQuery)
        
    If Not oRs.EOF Then
        oFindByQuery = IIf(IsNull(oRs(0).value), "", oRs(0).value)
    Else
        oFindByQuery = ""
    End If
    oCon.Close
    Exit Function
errhandler:
    MainModule.ShowMessage Err.Description, "FindData"
End Function




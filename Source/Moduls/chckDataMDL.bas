Attribute VB_Name = "chckDataMDL"
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
Public Function CheckData(sData As String, sFieldKriteria As String, sTable As String, sField As String) As Boolean
On Error GoTo errhandler

    If oCon.State = 1 Then oCon.Close
    oCon.Open MainModule.Conectionku(DBaseConection.Modul)
    sQuery = "Select " & sField & " from " & sTable & " where " & sFieldKriteria & " ='" & sData & "'"
    Set oRs = oCon.Execute(sQuery)
    If Not oRs.EOF Then
       CheckData = True
    Else
      CheckData = False
    End If
    oCon.Close
    Exit Function
    
errhandler:
    MainModule.ShowMessage Err.Description, "CheckData"
End Function

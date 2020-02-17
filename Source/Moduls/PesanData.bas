Attribute VB_Name = "PesanData"

Option Explicit
Dim cn As New ADODB.Connection
Dim rsPesan As New ADODB.Recordset
Dim nPesan As Integer

Public Sub PesanDataku(nField As String, nTable As String, nKondisi As String)
Dim txtQuery, cs   As String
cs = "driver={sql server};server=mtgmis22;initial catalog=Payroll;uid=sa;pwd=admin"
If cn.State = adStateOpen Then Set cn = Nothing
cn.Open cs
cn.CursorLocation = adUseClient

If rsPesan.State = adStateOpen Then rsPesan.Close
rsPesan.CursorLocation = adUseClient
rsPesan.CursorType = adOpenDynamic

txtQuery = " Select " & nField & " From " & nTable & " Where " & nKondisi
rsPesan.Open txtQuery, cn, adOpenDynamic, adLockOptimistic
If rsPesan.RecordCount >= 1 Then
nPesan = IIf(MsgBox("Data Sudah Ada", vbOKOnly) = vbOK, 1, 0)
End If

End Sub

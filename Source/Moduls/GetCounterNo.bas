Attribute VB_Name = "GetCounterNo"
Public Enum FlagTrans
    ODPSOrder = 1
    ODPSDelivery = 2
End Enum
Public Function GetNo(sFlagTrans As FlagTrans) As Double
Dim oCon As New ADODB.Connection
Dim oRs As New ADODB.Recordset
Dim sQuery As String
If oCon.State = 1 Then oCon.Close
oCon.Open MainModule.Conectionku(DBaseConection.Modul)

Select Case sFlagTrans
Case FlagTrans.ODPSOrder
sQuery = "Select top 1 OrderNumber from ODPSOrder_Dlvry Order By OrderNumber Desc"
Case FlagTrans.ODPSDelivery
sQuery = "Select top 1 ConfirmNumber from ODPSConfirm_Order Order By ConfirmNumber Desc"
End Select

Set oRs = oCon.Execute(sQuery)
If oRs.EOF Then
GetNo = 1
Else
GetNo = oRs(0) + 1
End If
oCon.Close

End Function

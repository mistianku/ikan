Attribute VB_Name = "PanahKnnKr"

Public Sub ShowBottom(txtForm As String)
Dim a As Integer

Dim kunci As String
kunci = "1"

For a = 0 To MenuFrm.cmbMenuku.ListCount
    If txtForm = MenuFrm.cmbMenuku.List(a) Then
       a = MenuFrm.cmbMenuku.ListCount
       kunci = "0"
    Else
    End If
      
Next
If kunci = "1" Then
   MenuFrm.cmbMenuku.AddItem txtForm
End If
If MenuFrm.cmbMenuku.ListCount = 0 Then
   MenuFrm.Toolbar1.Buttons(13).Enabled = False
   MenuFrm.Toolbar1.Buttons(14).Enabled = False
Else
    If MenuFrm.cmbMenuku.ListCount = 2 Then
       MenuFrm.Toolbar1.Buttons(13).Enabled = True
       MenuFrm.Toolbar1.Buttons(14).Enabled = False
    Else
        
    End If
End If
End Sub

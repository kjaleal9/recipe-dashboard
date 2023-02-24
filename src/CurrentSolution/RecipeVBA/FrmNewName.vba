
Private result As Integer
Private Sub CmdOk_Click()
If (TxtNewName.Text = "") Then
        MsgBox "Enter a name"
        Exit Sub
    Else
        DialogResult = vbOK
        Me.Hide
    End If
End Sub
Public Property Get DialogResult() As Integer
    DialogResult = result
End Property

Public Property Let DialogResult(ByVal vNewValue As Integer)
    result = vNewValue
End Property

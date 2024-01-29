Attribute VB_Name = "Module1"
Public c As ADODB.Connection
Public r As ADODB.Recordset
Public sql As String
Public Function Abc()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=prj2333e/prj2333e;Persist Security Info=False"
Set r = New ADODB.Recordset
End Function
Public Sub B()
RES = MsgBox("Do you Want to Exit", vbQuestion + vbYesNoCancel, "For Exit")
If (RES = vbYes) Then
End
End If

End Sub

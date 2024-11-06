
Option Explicit
Dim a As String

Private Sub Command2_Click()
Dim b As String
b = "Hola Mundo, como estan?"
Label1.Caption = b
End Sub

Private Sub Command3_Click()
a = "hola como estan"
Text2.Text = a
End Sub

Private Sub Form_Load()
' este es un comentario

End Sub

Private Sub Command1_Click()
a = "235"
Text1.Text = a
End Sub

Private Sub txtNombre_DblClick()
txtNombre.Text = "efrain"
End Sub

Private Sub Text2_Change()

End Sub

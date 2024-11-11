Option Explicit

Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    Dim X
    If Len(Trim(txtPassword.Text)) < 1 Then
        
        Select Case Cidioma
        Case 1: X = M_Pregunta(Caption, "No ha introducido un usuario correcto.", "Entrar como usuario genérico", "Salir de la aplicación")
        Case 2: X = M_Pregunta(Caption, "It has introduced a correct user.", "Login as user generic", "Exit the application")
        End Select
        
        
        If X = 2 Then
            Salir_Ya
        Else
            id_Usuario = ""
        End If
    End If
    Me.Hide
End Sub
Private Sub cmdCancel_GotFocus()
    Select Case Cidioma
    Case 1: M_Fijo "Cancelar proceso"
    Case 2: M_Fijo "Cancel Process"
    End Select
    
End Sub

Private Sub cmdOK_Click()
    'comprueba la contraseña correcta
    Dim I As Integer, X As String, sw As Boolean
    On Error Resume Next
    If UCase(Me.ActiveControl.Name) = UCase("txtUserName") Then
        If Len(Trim(txtPassword.Text)) < 1 Then
            txtPassword.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo 0
    txtUserName.Text = Trim(UCase(txtUserName.Text))
    txtPassword.Text = Trim(UCase(txtPassword.Text))
    id_User = txtUserName.Text
    Set tb1.ActiveConnection = dBa
    tb1.Source = "SELECT MATRICULA,NOMBRE FROM LT_EMPLEADOS WHERE " & _
          "USRNT=" & StrSql(id_User) & " AND USRCLAVE=" & StrSql(txtPassword.Text)
    tb1.Open , , adOpenForwardOnly, adLockReadOnly, adCmdText
    If tb1.EOF Then
        id_Matricula = ""
        id_Usuario = ""
    Else
        id_Matricula = Da_campo(tb1!MATRICULA)
        id_Usuario = Da_campo(tb1!Nombre)
    End If
    tb1.Close
    tb1.Source = "SELECT OPCION FROM LT_EMPLEACCESOS WHERE " & _
                 " MATRICULA=" & StrSql(id_Matricula) & " AND OPCION=" & StrSql("0000")
    tb1.Open
    sw = tb1.EOF
    tb1.Close
    If Not sw Then
        CargaTablaAcceso    ' Para quitar accesos de usuario anterior
        CargaTablaAccesoUsuario ' Ahora pongo los nuevos
        Me.Hide
        Exit Sub
    End If
    
    Select Case Cidioma
    Case 1: xx1 = "Usuario/contraseña no válido, vuelva a intentarlo"
    Case 2: xx1 = "User / password invalid retry"
    End Select
    MsgBox xx1, vbCritical + vbOKOnly, Caption
    
    txtPassword.Text = ""
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
End Sub
Private Sub cmdOK_GotFocus()
    Select Case Cidioma
    Case 1: M_Fijo "Aceptar Usuario y Clave de acceso."
    Case 2: M_Fijo "OK, and User Access Key."
    End Select
    
End Sub

Private Sub Form_Load()
    Dim I As Integer, X As String, aa As Variant
    txtUserName.Text = id_User
    txtPassword.Text = ""
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    
    Select Case Cidioma
    Case 1: M_Fijo "Introducir clave de acceso."
    Case 2: M_Fijo "Enter password."
    End Select
    
End Sub
Private Sub txtUserName_GotFocus()
    Select Case Cidioma
    Case 1: M_Fijo "Introducir usuario de acceso."
    Case 2: M_Fijo "Enter user access."
    End Select
    
End Sub

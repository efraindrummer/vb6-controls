Private Sub Form_Load()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim query As String
    
    ' Definir la cadena de conexión
    connString = "Provider=SQLOLEDB;Data Source=nombre_servidor;Initial Catalog=nombre_bd;User ID=usuario;Password=contraseña;"
    
    ' Definir la consulta SQL
    query = "SELECT * FROM nombre_tabla"
    
    ' Crear y abrir la conexión
    Set conn = New ADODB.Connection
    conn.Open connString
    
    ' Crear y abrir el recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    ' Enlazar el recordset al DataGrid
    Set DataGrid1.DataSource = rs
End Sub
' EJEMPLO CON CLICK EVENT
Private Sub cmdCargarDatos_Click()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim query As String
    
    ' Definir la cadena de conexión
    connString = "Provider=SQLOLEDB;Data Source=nombre_servidor;Initial Catalog=nombre_bd;User ID=usuario;Password=contraseña;"
    
    ' Definir la consulta SQL
    query = "SELECT * FROM nombre_tabla"
    
    ' Crear y abrir la conexión
    Set conn = New ADODB.Connection
    conn.Open connString
    
    ' Crear y abrir el recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    ' Enlazar el recordset al DataGrid
    Set DataGrid1.DataSource = rs
End Sub

' Cargar datos
Private Sub cmdCargarDatos_Click()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim connString As String
    Dim query As String
    
    ' Definir la cadena de conexión
    connString = "Provider=SQLOLEDB;Data Source=nombre_servidor;Initial Catalog=nombre_bd;User ID=usuario;Password=contraseña;"
    
    ' Definir la consulta SQL para cargar datos
    query = "SELECT * FROM nombre_tabla"
    
    ' Crear y abrir la conexión
    Set conn = New ADODB.Connection
    conn.Open connString
    
    ' Crear y abrir el recordset
    Set rs = New ADODB.Recordset
    rs.Open query, conn, adOpenStatic, adLockReadOnly
    
    ' Enlazar el recordset al DataGrid
    Set DataGrid1.DataSource = rs
    
    ' Cerrar la conexión
    Set conn = Nothing
    Set rs = Nothing
End Sub
' INSERTAR DATOS
Private Sub cmdInsertarDatos_Click()
    Dim conn As ADODB.Connection
    Dim insertQuery As String
    Dim connString As String
    Dim valorColumna1 As String
    Dim valorColumna2 As String
    ' Suponiendo que vamos a insertar valores de la primera y segunda columna seleccionada
    
    ' Verificar que hay una fila seleccionada
    If DataGrid1.Row < 0 Then
        MsgBox "Por favor, selecciona una fila para insertar."
        Exit Sub
    End If
    
    ' Obtener los valores de las columnas deseadas (por ejemplo, la primera y la segunda columna)
    valorColumna1 = DataGrid1.Columns(0).Text
    valorColumna2 = DataGrid1.Columns(1).Text

    ' Definir la cadena de conexión
    connString = "Provider=SQLOLEDB;Data Source=nombre_servidor;Initial Catalog=nombre_bd;User ID=usuario;Password=contraseña;"
    
    ' Crear la consulta de inserción
    insertQuery = "INSERT INTO otra_tabla (columna1, columna2) VALUES ('" & valorColumna1 & "', '" & valorColumna2 & "')"
    
    ' Crear y abrir la conexión
    Set conn = New ADODB.Connection
    conn.Open connString
    
    ' Ejecutar la consulta de inserción
    conn.Execute insertQuery
    
    ' Confirmar la inserción
    MsgBox "Datos insertados correctamente en otra tabla."
    
    ' Cerrar la conexión
    conn.Close
    Set conn = Nothing
End Sub

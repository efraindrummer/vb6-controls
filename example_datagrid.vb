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

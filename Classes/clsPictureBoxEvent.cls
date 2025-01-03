VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPictureBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents picBox As PictureBox
Attribute picBox.VB_VarHelpID = -1

Private Sub picBox_DblClick()
    MsgBox "Cargando detalle..."
    Dim metaTitle As String
    Dim query As String
    
    ' Obtener el t�tulo de la meta desde el Tag
    metaTitle = picBox.Tag
        
    ' Crear y mostrar una nueva instancia del formulario frmDetalle
    Dim frmDetalle As frmDetalle
    Set frmDetalle = New frmDetalle
    
    ' Crear la conexi�n a SQL Server
    Dim conn As ADODB.connection
    Dim rs As ADODB.Recordset
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
    conn.Open
    
    ' Escapar comillas simples en el valor del t�tulo
    metaTitle = Replace(metaTitle, "'", "''")
    
    ' Construir la consulta SQL
    query = "SELECT * FROM Goals WHERE title = '" & metaTitle & "'"
    
    ' Ejecutar la consulta
    Set rs = conn.Execute(query)
    
        ' Obtener el valor del campo "completionMonth"
    Dim monthIndex As Integer
    monthIndex = rs.Fields("completionMonth").Value

    
    ' Verificar si hay resultados
    If Not rs.EOF Then
        ' Asignar informaci�n relevante al formulario
        frmDetalle.txtTituloMeta.Text = rs.Fields("title").Value
        frmDetalle.txtDescripcionMeta.Text = rs.Fields("description").Value
        
                ' Validar que el �ndice est� dentro del rango v�lido
        If monthIndex >= 0 And monthIndex <= 11 Then
            ' Restar 1 para ajustar al �ndice 0-basado del ComboBox
            frmDetalle.cmbFechaTentativaMeta.ListIndex = monthIndex '- 1
        Else
            MsgBox "El valor de completionMonth no es v�lido: " & monthIndex, vbExclamation, "Error de datos"
        End If
        
        ' Cargar la imagen en el PictureBox del formulario
        Dim imagePath As String
        imagePath = rs.Fields("imageUrl").Value
        
        If Len(Dir(imagePath)) > 0 Then
            frmDetalle.picMetaImage.Picture = LoadPicture(imagePath)
        Else
            MsgBox "No se encontr� la imagen en la ruta especificada: " & imagePath, vbExclamation, "Error de carga"
        End If
    Else
        MsgBox "No se encontr� informaci�n para la meta seleccionada.", vbInformation, "Sin datos"
    End If
    
    ' Mostrar el formulario
    frmDetalle.Show
    
    ' Limpiar recursos
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


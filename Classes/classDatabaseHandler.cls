VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classDatabaseHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public connection As ADODB.connection

'Public Function ConnectDb() As Boolean
'    On Error GoTo ERR_HANDLER
    
'    Set global_conection = New ADODB.connection
'    global_connection.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
'    global_connection.Open
    
'    ConnectDb = True
    
'ERR_HANDLER:
'    ConnectDb = False

'End Function


'Public Function InsertarElemento(ByVal strRutaImagen As String, ByVal strTitulo As String, ByVal strDescripcion As String, ByVal strFechaTentativa As String) As Boolean
    
'   Dim strSQL As String
    
'    strSQL = "INSERT INTO Goals (title, description, completionMonth, imageUrl) VALUES ('" & Replace(strRutaImagen, "'", "''") & "', '" & Replace(strTitulo, "'", "''") & "', '" & Replace(strDescripcion, "'", "''") & "')"

'    global_connection.Execute strSQL, , adCmdText
'    InsertarElemento = True
'    Exit Function

'End Function

'Public Function obtenerElementos() As ADODB.Recordset
'    Dim registros As ADODB.Recordset
'    Set registros = New ADODB.Recordset
    
'    If global_connection Is Nothing Or global_connection.State <> adStateOpen Then
'        MsgBox "NO abierta", vbCritical, "Error"
'        Exit Function
'    End If
    
'    registros.Open "SELECT IdGoal, title, description, completionMonth, imageUrl FROM Goals", global_connection, adOpenForwardOnly, adLockReadOnly
    
'    If Not registros.EOF Then
'        Dim itmX As ListItem
'        Do Until registros.EOF
'            registros.MoveNext
'        Loop
'    End If
    
'    obtenerElementos = registros
    
'    Exit Function
    
'End Function

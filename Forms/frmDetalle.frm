VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDetalle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Agregar Nueva Meta"
   ClientHeight    =   9105
   ClientLeft      =   11655
   ClientTop       =   7680
   ClientWidth     =   14640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   14640
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog commonDialogInsertarImagen 
      Left            =   840
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picMetaImage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   6960
      ScaleHeight     =   5000
      ScaleMode       =   0  'User
      ScaleWidth      =   5000
      TabIndex        =   11
      Top             =   1680
      Width           =   5055
   End
   Begin VB.Frame metaDetalles 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Detalles"
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   5655
      Begin VB.CommandButton cmdBtnGuardarMeta 
         Appearance      =   0  'Flat
         Caption         =   "Guardar meta"
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   4200
         Width           =   4455
      End
      Begin VB.CommandButton cmdBtnImagen 
         Appearance      =   0  'Flat
         Caption         =   "Insertar Imagen"
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox cmbFechaTentativaMeta 
         Height          =   330
         ItemData        =   "frmDetalle.frx":0000
         Left            =   360
         List            =   "frmDetalle.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txtDescripcionMeta 
         Height          =   975
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Descripción de tu meta"
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox txtTituloMeta 
         Height          =   495
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Título meta"
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblImagenRepresentativa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Imagen representativa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblFechaTentativa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fecha tentativa de alcanzarla"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblDescripcionMeta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Descripción"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblTituloMeta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Titulo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Label lblTituloDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Detalle de Meta"
      BeginProperty Font 
         Name            =   "Roboto Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.connection
' declare values
Dim titulo As String
Dim descripcion As String
Dim fechaCom As String
Dim urlImagen As String

Private Sub cmdBtnGuardarMeta_Click()

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Call ConnectToDb
    
        ' Ensure conn is initialized before checking its state
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
        
            titulo = txtTituloMeta.Text
            descripcion = txtDescripcionMeta.Text
            fechaCom = cmbFechaTentativaMeta.ListIndex ' TENTATIVOOOO
            'urlImagen = ' es que es rutaImagen global
            
            On Error GoTo ErrorHandler
            'rs.Open "SELECT * FROM Goals", conn, adOpenStatic, adLockReadOnly

            ' Create the INSERT statement with proper concatenation
            Dim sql As String
            
            sql = "INSERT INTO Goals (title, description, completionMonth, imageUrl) VALUES (" & _
                  "'" & Replace(titulo, "'", "''") & "', " & _
                  "'" & Replace(descripcion, "'", "''") & "', " & _
                  fechaCom & ", " & _
                  "'" & Replace(rutaImagen, "'", "''") & "')"
            
            ' Execute the SQL query
            conn.Execute sql
            
            MsgBox "Meta guardada correctamente.", vbInformation
            
            Call DisconnectToDb ' Close my conn
            Exit Sub
        End If
    Else
        MsgBox "Connection object not initialized.", vbCritical
    End If

ErrorHandler:
    MsgBox "Errorrrrr: " & Err.description, vbCritical
    Exit Sub
    
    'If ConnectDb = True Then
    '    Debug.Print "HOLAAA"
    'Else
    '    MsgBox "Error en DB: " & Err.Description, vbCritical, "Error"
    'End If

End Sub

Private Sub cmdBtnImagen_Click()

    On Error GoTo ERR_HANDLER

    With commonDialogInsertarImagen ' Component MSCommon Dialog 6.0
        .DialogTitle = "Selecciona imagen representativa"
        .Filter = "Archivos de imagen|*.jpg;*.png"
        '.Flags = cdlOFNAllowMultiselect ' permite seleccion multiple de archivos
        .ShowOpen
        rutaImagen = .FileName
    End With
    
    If rutaImagen <> "" Then
        Dim objLoader As classImageLoader
        Set objLoader = New classImageLoader
        
        'CargarImagen rutaImagen
        Dim myImage As StdPicture
        Set myImage = objLoader.CargarImagen(rutaImagen, picMetaImage)
        
    End If
    
    Exit Sub
    
ERR_HANDLER:
        Debug.Print "Ocurrió error. " & Err.description
        MsgBox "Error en carga de imagen: " & Err.description, vbCritical, "Error"
    
End Sub


Private Sub Form_Load()
    
    ' agregando elementos a combobox de fecha tentativa de alcanzar meta
    cmbFechaTentativaMeta.AddItem "Enero"
    cmbFechaTentativaMeta.AddItem "Febrero"
    cmbFechaTentativaMeta.AddItem "Marzo"
    cmbFechaTentativaMeta.AddItem "Abril"
    cmbFechaTentativaMeta.AddItem "Mayo"
    cmbFechaTentativaMeta.AddItem "Junio"
    cmbFechaTentativaMeta.AddItem "Julio"
    cmbFechaTentativaMeta.AddItem "Agosto"
    cmbFechaTentativaMeta.AddItem "Septiembre"
    cmbFechaTentativaMeta.AddItem "Octubre"
    cmbFechaTentativaMeta.AddItem "Noviembre"
    cmbFechaTentativaMeta.AddItem "Diciembre"

End Sub

' Conexion de base de datos DB Connection
Public Sub ConnectToDb()
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
    conn.Open
End Sub

Public Sub DisconnectToDb()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub


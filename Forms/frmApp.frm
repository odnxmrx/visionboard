VERSION 5.00
Begin VB.MDIForm frmApp 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Vision Board"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   13845
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu smNuevo 
         Caption         =   "Crear Nuevo"
      End
      Begin VB.Menu mSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mAcercaDe 
      Caption         =   "Acerca de"
      Begin VB.Menu smAcercaDe 
         Caption         =   "Créditos"
      End
   End
End
Attribute VB_Name = "frmApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mSalir_Click()
    Unload Me
End Sub

Private Sub smNuevo_Click()
    Set frmDetalle = New frmDetalle ' Nueva instancia de Detalle
    frmDetalle.Show
End Sub

' Conexion de base de datos DB Connection
Public Sub ConnectToDb()
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
    conn.Open
End Sub

Public Sub DisconnectToDb()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub


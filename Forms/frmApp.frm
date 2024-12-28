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
   Begin VB.PictureBox pictureBoxMain 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   0
      Picture         =   "frmApp.frx":0000
      ScaleHeight     =   8265
      ScaleMode       =   0  'User
      ScaleWidth      =   13815
      TabIndex        =   0
      Top             =   0
      Width           =   13845
   End
   Begin VB.Menu mArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu smNuevo 
         Caption         =   "Crear Nuevo"
      End
      Begin VB.Menu mVerVisionboard 
         Caption         =   "Ver Visionboard"
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
Public conn As ADODB.connection

Private Sub mSalir_Click()
    Unload Me
End Sub

Private Sub mVerVisionboard_Click()
    pictureBoxMain.Visible = False
    frmVisionboard.Show
End Sub

Private Sub smAcercaDe_Click()
    MsgBox "Visionboard project" & vbCrLf & _
    "" & vbCrLf & _
    "Developed on Visual Basic 6," & vbCrLf & _
    "Microsoft SQL Server" & vbCrLf & _
    "" & vbCrLf & _
    "Desarrollado por:" & vbCrLf & _
    "Carlos Aldair Ortiz Mata" & vbCrLf & _
    "Armando Pineda Gama" & vbCrLf & _
    "© Todos los derechos reservados. Diciembre, 2024.", vbOKOnly, "Acerca de"
End Sub

Private Sub smNuevo_Click()
    pictureBoxMain.Visible = False
    Set frmDetalle = New frmDetalle ' Nueva instancia de Detalle
    frmDetalle.Show
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


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

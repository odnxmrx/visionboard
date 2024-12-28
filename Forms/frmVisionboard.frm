VERSION 5.00
Begin VB.Form frmVisionboard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   13050
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCargarVisionboard 
      Appearance      =   0  'Flat
      Caption         =   "Mostrar visionboard"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdBtnSalir 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   375
      Left            =   10680
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmVisionboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.connection
Dim rs As ADODB.Recordset
Dim imgData() As Byte
Dim randomIndex As Integer
Dim query As String
Dim totalImagenes As Integer

Private Sub cmdBtnSalir_Click()
    Unload Me
End Sub

Private Sub cmdCargarVisionboard_Click()
    ' Crear la conexión a SQL Server
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
    conn.Open

    ' Obtener el número total de imágenes en la base de datos
    Set rs = conn.Execute("SELECT COUNT(*) AS TotalImagenes FROM Goals")
    totalImagenes = rs.Fields("TotalImagenes").Value

    ' Configurar variables para la disposición de los PictureBox
    Dim imgTop As Single
    Dim imgLeft As Single
    Dim imgWidth As Single
    Dim imgHeight As Single
    Dim spacing As Single
    Dim i As Integer

    ' Inicializar propiedades
    imgTop = 1000
    imgLeft = 10
    imgWidth = 3000 ' Ancho del PictureBox
    imgHeight = 3000 ' Alto del PictureBox
    spacing = 50 ' Espaciado entre PictureBox

    ' Iterar para crear y cargar imágenes en los PictureBox
    For i = 1 To totalImagenes
        ' Obtener la ruta de cada imagen y los datos asociados
        query = "SELECT IdGoal, title, description, completionMonth, imageUrl FROM Goals WHERE IdGoal = " & i
        Set rs = conn.Execute(query)

        Dim imgPath As String
        Dim goalId As Integer
        Dim title As String
        Dim description As String
        Dim completionMonth As String

        goalId = rs.Fields("IdGoal").Value
        title = rs.Fields("title").Value
        description = rs.Fields("description").Value
        completionMonth = rs.Fields("completionMonth").Value
        imgPath = rs.Fields("imageUrl").Value

        ' Crear el PictureBox dinámicamente
        Dim picBox As PictureBox
        Set picBox = Me.Controls.Add("VB.PictureBox", "picBox" & i)
        With picBox
            .Width = imgWidth
            .Height = imgHeight
            .Left = imgLeft
            .Top = imgTop
            .Visible = True
            .AutoSize = False
            .Appearance = 0
            .BorderStyle = vbFixedSingle

            ' Almacenar datos como propiedades del PictureBox
            .Tag = goalId & "|" & title & "|" & description & "|" & completionMonth & "|" & imgPath

            ' Verificar si la ruta de la imagen es válida antes de cargarla
            If Len(Dir(imgPath)) > 0 Then
                Dim objLoader As classImageLoader
                Set objLoader = New classImageLoader
                objLoader.CargarImagen imgPath, picBox
                Set objLoader = Nothing
            Else
                MsgBox "La ruta de la imagen no es válida: " & imgPath, vbExclamation, "Error de Carga"
            End If
        End With

        ' Asignar evento Click dinámico al PictureBox
        'picBox.OnClick = "[EventProcedure]"

        ' Actualizar posición para el siguiente PictureBox
        imgLeft = imgLeft + imgWidth + spacing
        If imgLeft + imgWidth > Me.ScaleWidth Then
            imgLeft = 10
            imgTop = imgTop + imgHeight + spacing
        End If
    Next i

    ' Cerrar la conexión
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


Private Sub picBox_Click(Index As Integer)
    ' Obtener el control que disparó el evento
    Dim clickedPicBox As PictureBox
    Set clickedPicBox = Me.Controls("picBox" & Index)

    ' Extraer los datos almacenados en el Tag
    Dim data() As String
    data = Split(clickedPicBox.Tag, "|")

    ' Abrir el formulario frmDetalle y precargar los datos
    Dim frm As New frmDetalle
    With frm
        .txtTituloMeta.Text = data(1) ' Título
        .txtDescripcionMeta.Text = data(2) ' Descripción
        .cmbFechaTentativaMeta.ListIndex = CInt(data(3)) ' Mes de finalización
        Dim objLoader As classImageLoader
        Set objLoader = New classImageLoader
        objLoader.CargarImagen data(4), .picMetaImage ' Imagen
        Set objLoader = Nothing

        ' Mostrar el formulario
        .Show
    End With
End Sub





Private Sub Form_Unload(Cancel As Integer)
    frmApp.pictureBoxMain.Visible = True ' Volver a mostrar imagen de fondo en MDF Principal
End Sub


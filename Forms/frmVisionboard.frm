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
Private picBoxArray() As clsPictureBoxEvent ' Arreglo para manejar eventos

Private Sub cmdBtnSalir_Click()
    Unload Me
End Sub


Public Sub cmdCargarVisionboard_Click()
    ' Crear la conexi�n a SQL Server
    Set conn = New ADODB.connection
    conn.ConnectionString = "Provider=SQLOLEDB; Data Source=.; Initial Catalog=Visionboarddb; Trusted_Connection=Yes;"
    conn.Open

    ' Obtener el n�mero total de im�genes en la base de datos
    Set rs = conn.Execute("SELECT COUNT(*) AS TotalImagenes FROM Goals")
    totalImagenes = rs.Fields("TotalImagenes").Value

    ' Configurar variables para la disposici�n de los PictureBox
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

    ' Redimensionar el arreglo de clases
    ReDim picBoxArray(1 To totalImagenes)

    ' Iterar para crear y cargar im�genes en los PictureBox
    For i = 1 To totalImagenes
        ' Obtener datos de la meta
        query = "SELECT IdGoal, title, imageUrl FROM Goals WHERE IdGoal = " & i
        Set rs = conn.Execute(query)

        ' Crear din�micamente un PictureBox
        Dim newPicBox As PictureBox
        Set newPicBox = Me.Controls.Add("VB.PictureBox", "picBox" & i)

        ' Configurar el PictureBox
        With newPicBox
            .Width = imgWidth
            .Height = imgHeight
            .Left = imgLeft
            .Top = imgTop
            .Visible = True
            .AutoSize = False
            .Appearance = 0
            .BorderStyle = vbFixedSingle
            .Tag = rs.Fields("title").Value ' Guardar ID en el Tag
        End With

        ' Vincular el PictureBox a la clase para manejar el evento DblClick
        Set picBoxArray(i) = New clsPictureBoxEvent
        Set picBoxArray(i).picBox = newPicBox

        ' Cargar la imagen si la ruta es v�lida
        If Len(Dir(rs.Fields("imageUrl").Value)) > 0 Then
            Dim objLoader As classImageLoader
            Set objLoader = New classImageLoader
            objLoader.CargarImagen rs.Fields("imageUrl").Value, newPicBox
            Set objLoader = Nothing
        End If

        ' Actualizar posici�n para el siguiente PictureBox
        imgLeft = imgLeft + imgWidth + spacing
        If imgLeft + imgWidth > Me.ScaleWidth Then
            imgLeft = 10
            imgTop = imgTop + imgHeight + spacing
        End If
    Next i

    ' Cerrar la conexi�n
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub


Private Sub picBox_Click(Index As Integer)
    ' Obtener el control que dispar� el evento
    Dim clickedPicBox As PictureBox
    Set clickedPicBox = Me.Controls("picBox" & Index)

    ' Extraer los datos almacenados en el Tag
    Dim data() As String
    data = Split(clickedPicBox.Tag, "|")

    ' Abrir el formulario frmDetalle y precargar los datos
    Dim frm As New frmDetalle
    With frm
        .txtTituloMeta.Text = data(1) ' T�tulo
        .txtDescripcionMeta.Text = data(2) ' Descripci�n
        .cmbFechaTentativaMeta.ListIndex = CInt(data(3)) ' Mes de finalizaci�n
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


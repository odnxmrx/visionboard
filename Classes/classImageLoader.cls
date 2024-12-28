VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classImageLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Function que hace la carga de la imagen
Public Function CargarImagen(ByVal ruta As String, picMetaImage As PictureBox) As StdPicture
    On Error GoTo ERR_HANDLER ' Handle runtime errors gracefully

    ' Load the picture
    picMetaImage.Picture = LoadPicture(ruta)

    ' Call the resizing logic
    RedimensionarImagen picMetaImage

    Exit Function
ERR_HANDLER:
    MsgBox "Ocurrió errorsito en CargarImagen Fn" & Err.Description, vbCritical, "Error"
    Resume Next
End Function

' Subroutine para redimensionar la imagen dentro de un PictureBox
Public Sub RedimensionarImagen(picMetaImage As PictureBox)
    Dim originalWidth As Single
    Dim originalHeight As Single
    Dim scaleX As Single
    Dim scaleY As Single
    Dim targetWidth As Single
    Dim targetHeight As Single
    Dim xOffset As Single
    Dim yOffset As Single

    ' Get the original image dimensions
    originalWidth = picMetaImage.Picture.Width / Screen.TwipsPerPixelX
    originalHeight = picMetaImage.Picture.Height / Screen.TwipsPerPixelY

    ' Calculate scale factors for resizing
    scaleX = picMetaImage.ScaleWidth / originalWidth
    scaleY = picMetaImage.ScaleHeight / originalHeight

    ' Use the smaller scale factor to fit the image inside the PictureBox
    If scaleX < scaleY Then
        targetWidth = originalWidth * scaleX
        targetHeight = originalHeight * scaleX
    Else
        targetWidth = originalWidth * scaleY
        targetHeight = originalHeight * scaleY
    End If

    ' Center the image inside the PictureBox
    xOffset = (picMetaImage.ScaleWidth - targetWidth) / 2
    yOffset = (picMetaImage.ScaleHeight - targetHeight) / 2

    ' Clear the PictureBox and draw the resized image
    picMetaImage.Cls
    picMetaImage.PaintPicture picMetaImage.Picture, xOffset, yOffset, targetWidth, targetHeight
End Sub


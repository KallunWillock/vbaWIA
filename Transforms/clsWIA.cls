VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
                                                                                                                                           ' _
  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                      ' _
  ||||||||||||||||||||||||||             clsWIA (v1.3)             ||||||||||||||||||||||||||||||||||                                      ' _
  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                           ' _
  AUTHOR:   Kallun Willock                                                                                                                                                                ' _
  URL:      https://www.mrexcel.com/board/threads/resize-image-inside-a-image-control-on-userform.1234401/                                 ' _
  PURPOSE:  A simple class module that leverages the functionality of the WIA COM Object for image manipulation/conversion.                ' _
            It was developed to respond to the specific requirements set out in the thread at the above-referenced URL.                    ' _
  LICENSE:  MIT                                                                                                                            ' _
  VERSION:  1.0        08/04/2023         Published v1.0 on Mr.Excel forum.                                                                                                               ' _
            1.1        18/05/2023         Added comments for publication Github                                                                                                           ' _
            1.2        26/05/2023         Added flip routine.                                                                                                                             ' _
            1.3        02/03/2026         Fixed bugs.                                                                                      ' _
  TODO:     [ ] Error Handling                                                                                                                                                            ' _

Option Explicit

Public Enum RotationEnum
    Rotation0 = 0
    Rotation90 = 90
    Rotation180 = 180
    Rotation270 = 270
End Enum

Public Enum ImageFormatEnum2
  bmp
  PNG
  GIF
  JPG
  TIFF
End Enum

Private Type ClassVariables
  FilePath          As String
  ImageFile         As Object
  DestinationPath   As String
  SaveAsFilePath    As String
End Type
Private This As ClassVariables

Public Property Let FilePath(ByVal RHS As String)
  
  ' The image will be loaded by passing the filename to the Source Property.
  If Len(VBA.Trim(RHS)) = 0 Then Exit Property
  On Error GoTo ErrHandler
  If (GetAttr(RHS) And vbDirectory) <> vbDirectory Then
    This.FilePath = RHS
    Set This.ImageFile = CreateObject("WIA.ImageFile")
    This.ImageFile.LoadFile RHS
  End If
ErrHandler:
End Property

Public Property Get FilePath() As String:             FilePath = This.FilePath:           End Property
Public Property Let SaveAs(ByVal RHS As String):      This.SaveAsFilePath = RHS:          End Property
Public Property Get SaveAs() As String:               SaveAs = This.SaveAsFilePath:       End Property

Public Property Get Width() As Long
  
  If Not This.ImageFile Is Nothing Then
    Width = This.ImageFile.Width
  End If

End Property
Public Property Get Height() As Long

  If Not This.ImageFile Is Nothing Then
    Height = This.ImageFile.Height
  End If

End Property

Public Property Get Picture() As stdole.StdPicture
  ' The Picture property converts the WIA ImageFile Object into a standard stdPicture object.
  
  If Not This.ImageFile Is Nothing Then
    Set Picture = This.ImageFile.FileData.Picture
  End If
  
End Property

Public Sub ScaleImage(ByVal NewWidth As Long, ByVal NewHeight As Long, _
                      Optional ByVal MaintainAspect As Boolean = True)
  
  If This.ImageFile Is Nothing Then Exit Sub
        
  If NewWidth > 0 And NewHeight > 0 Then
    With CreateObject("WIA.ImageProcess")
      .Filters.Add .FilterInfos("Scale").FilterID
      .Filters(1).Properties("MaximumHeight").value = NewHeight
      .Filters(1).Properties("MaximumWidth").value = NewWidth
      .Filters(1).Properties("PreserveAspectRatio") = MaintainAspect
    Set This.ImageFile = .Apply(This.ImageFile)
    End With
  End If
    
End Sub

Public Sub RotateImage(Optional ByVal DegreesRotate As RotationEnum)
  ' Rotations are only available at 0, 90, 180, 270 degrees
  ' 0 degrees corresponds to the default orientation of the image (as loaded), and each
  ' sebsequent 90 degree rotation transforms clock-wise.
  
  If This.ImageFile Is Nothing Then Exit Sub
  
  With CreateObject("WIA.ImageProcess")
    .Filters.Add .FilterInfos("RotateFlip").FilterID
    .Filters(1).Properties("RotationAngle") = DegreesRotate
    Set This.ImageFile = .Apply(This.ImageFile)
  End With
  
End Sub

Public Sub FlipImage(ByVal Horizontal As Boolean, Optional ByVal Vertical As Boolean = False)

  If This.ImageFile Is Nothing Then Exit Sub
  
  With CreateObject("WIA.ImageProcess")
    .Filters.Add .FilterInfos("RotateFlip").FilterID
    .Filters(1).Properties("FlipHorizontal") = Horizontal
    .Filters(1).Properties("FlipVertical") = Vertical
  End With

End Sub

Public Function GetPixelColor(ByVal X As Long, ByVal Y As Long) As Long
  
  If This.ImageFile Is Nothing Then Exit Function
    
  Dim Location As Long, TargetColor As Long
  Location = X + (Y * This.ImageFile.Width)
  TargetColor = This.ImageFile.ARGBData(Location)
  GetPixelColor = ARGBtoRGB(TargetColor)

End Function

Public Function ARGBtoRGB(ByVal ARGB As Long) As Long
  
  Dim R As Byte, G As Byte, B As Byte
  ARGB = ARGB And &HFFFFFF
  R = ARGB And &HFF
  G = ((ARGB And &HFF00) / &H100) And &HFF
  B = (ARGB And &HFF0000) / &H10000
  ARGBtoRGB = RGB(B, G, R)

End Function

Public Sub SaveImage(Optional ByVal ImageFileType As ImageFormatEnum2)
  ' A straight-forward image file format conversion method
  
  Dim FormatFiltersArray  As Variant
  Dim FormatFilter        As String
  Dim Extension           As String

  FormatFiltersArray = Array("AB", "AF", "B0", "AE", "B1")
  FormatFilter = "{B96B3C" & FormatFiltersArray(ImageFileType) & "-0728-11D3-9D7B-0000F81EF32E}"

  If This.SaveAsFilePath = vbNullString Then
    Extension = Choose(ImageFileType + 1, "BMP", "PNG", "GIF", "JPG", "TIFF")
    This.SaveAsFilePath = Replace(This.FilePath, This.ImageFile.FileExtension, Extension, , , vbTextCompare)
  End If

  With CreateObject("WIA.ImageProcess")
    .Filters.Add .FilterInfos("Convert").FilterID
    .Filters(1).Properties("FormatID").value = FormatFilter
    .Filters(1).Properties("Quality").value = 100
    Set This.ImageFile = .Apply(This.ImageFile)
  End With

  ' If the target filename exists, it will be deleted.
  If Len(Dir(This.SaveAsFilePath)) Then Kill This.SaveAsFilePath
  
  This.ImageFile.SaveFile This.SaveAsFilePath
    
End Sub

Private Sub Class_Terminate()
  
  Set This.ImageFile = Nothing

End Sub

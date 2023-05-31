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
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||             clsWIA (v1.2)             ||||||||||||||||||||||||||||||||||                                                                                     ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                        ' _
AUTHOR:   Kallun Willock                                                                                                                                                                ' _
URL:      https://www.mrexcel.com/board/threads/resize-image-inside-a-image-control-on-userform.1234401/                                                                                ' _
PURPOSE:  A simple class module that leverages the functionality of the WIA COM Object for image manipulation/conversion.                                                               ' _
          It was developed to respond to the specific requirements set out in the thread at the above-referenced URL.                                                                   ' _
LICENSE:  MIT                                                                                                                                                                           ' _
VERSION:  1.0        08/04/2023         Published v1.0 on Mr.Excel forum.                                                                                                               ' _
          1.1        18/05/2023         Added comments for publication Github                                                                                                           ' _
          1.2        26/05/2023         Added flip routine.                                                                                                                             ' _
TODO:     [ ] Error Handling                                                                                                                                                            ' _

Option Explicit

Public Enum RotationEnum
    Rotation0 = 0
    Rotation90 = 90
    Rotation180 = 180
    Rotation270 = 270
End Enum

Public Enum ImageFormatEnum
    BMP
    PNG
    GIF
    JPG
    TIFF
End Enum

Private TargetImageFile     As Object
Private TargetImageProcess  As Object
Private cSource             As String
Private cDestination        As String

Public Property Let Source(ByVal uSource As String)
    ' The image will be loaded by passing the filename to the Source Property.
    
    If uSource = vbNullString Then Exit Property
    If Len(Dir(uSource)) Then
        cSource = uSource
        Set TargetImageFile = CreateObject("WIA.ImageFile")
        TargetImageFile.LoadFile uSource
    End If
    
End Property

Public Property Get Source() As String:                         Source = cSource:               End Property

Public Property Let Destination(ByVal uDestination As String):  cDestination = uDestination:    End Property
Public Property Get Destination() As String:                    Destination = cDestination:     End Property

Public Property Get Picture() As stdole.StdPicture
    ' The Picture property converts the WIA ImageFile Object into a standard stdPicture object.
    
    If Not TargetImageFile Is Nothing Then
        Set Picture = TargetImageFile.FileData.Picture
    End If
    
End Property

Public Sub ScaleImage(ByVal NewWidth As Long, ByVal NewHeight As Long, _
                      Optional ByVal MaintainAspect As Boolean = True)

    If TargetImageFile Is Nothing Then Exit Sub
                      
    If NewWidth > 0 And NewHeight > 0 Then
        With CreateObject("WIA.ImageProcess")
            .Filters.Add .FilterInfos("Scale").FilterID
            .Filters(1).Properties("MaximumHeight").Value = NewHeight
            .Filters(1).Properties("MaximumWidth").Value = NewWidth
            .Filters(1).Properties("PreserveAspectRatio") = MaintainAspect
            Set TargetImageFile = .Apply(TargetImageFile)
        End With
    End If
    
End Sub

Public Sub RotateImage(Optional ByVal DegreesRotate As RotationEnum)
    ' Rotations are only available at 0, 90, 180, 270 degrees
    ' 0 degrees corresponds to the default orientation of the image (as loaded), and each
    ' sebsequent 90 degree rotation transforms clock-wise.

    If TargetImageFile Is Nothing Then Exit Sub
    
    With CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("RotateFlip").FilterID
        .Filters(1).Properties("RotationAngle") = DegreesRotate
        Set TargetImageFile = .Apply(TargetImageFile)
    End With

End Sub

Public Sub FlipImage(ByVal Horizontal As Boolean, Optional ByVal Vertical As Boolean = False)

    If TargetImageFile Is Nothing Then Exit Sub

    With CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("RotateFlip").FilterID
        .Filters(1).Properties("FlipHorizontal") = Horizontal
        .Filters(1).Properties("FlipVertical") = Vertical
    End With

End Sub

Public Sub SaveImage(Optional ByVal ImageFileType As ImageFormatEnum)
    ' A straight-forward image file format conversion method
    
    Dim FormatFiltersArray  As Variant
    Dim FormatFilter        As String
    Dim Extension           As String
  
    FormatFiltersArray = Array("AB", "AF", "B0", "AE", "B1")
    FormatFilter = "{B96B3C" & FormatFiltersArray(ImageFileType) & "-0728-11D3-9D7B-0000F81EF32E}"
  
    If cDestination = vbNullString Then
        Extension = Choose(ImageFileType + 1, "BMP", "PNG", "GIF", "JPG", "TIFF")
        cDestination = Replace(cSource, TargetImageFile.FileExtension, Extension, , , vbTextCompare)
    End If
  
    With CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("Convert").FilterID
        .Filters(1).Properties("FormatID").Value = FormatFilter
        .Filters(1).Properties("Quality").Value = 100
        Set TargetImageFile = .Apply(TargetImageFile)
    End With
  
    ' If the target filename exists, it will be deleted.
    If Len(Dir(cDestination)) Then Kill cDestination
    
    TargetImageFile.SaveFile cDestination
    
End Sub

Private Sub Class_Terminate()
    Set TargetImageFile = Nothing
End Sub




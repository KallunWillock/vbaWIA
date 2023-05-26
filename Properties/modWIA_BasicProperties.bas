Attribute VB_Name = "modWIA_BasicProperties"
                                                                                                                                                                                        ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||      modWIA_BasicProperties (v1.1)    ||||||||||||||||||||||||||||||||||                                                                                     ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                        ' _
AUTHOR:   Kallun Willock                                                                                                                                                                ' _

URL:      https://www.mrexcel.com/board/threads/working-with-images-in-vba-image-properties.1224361/                                                                                    ' _

PURPOSE:  A simple class module that leverages the functionality of the WIA COM Object for image manipulation/conversion.                                                               ' _
          It was developed to respond to the specific requirements set out in the thread at the above-referenced URL.                                                                   ' _

LICENSE:  MIT                                                                                                                                                                           ' _

VERSION:  1.0        12/12/2022         Published v1.0 on Mr.Excel forum.                                                                                                               ' _
          1.1        26/05/2023         Added comments for publication Github                                                                                                           ' _

TODO:     [ ] Error Handling                                                                                                                                                            ' _

Option Explicit

' WIA - accessing image properties
'
' The following function (and test procedure) demonstrates
' how to access basic image properties with the WIA COM Object.
' I have used late binding here, but you can enable
' early-binding (and intellisense) by selecting:
'
' Tools > References > Microsoft Windows Image Acquisition Library
'
' Dan_W

Enum ImagePropertyTypeEnum
    ImageWidth
    ImageHeight
    HorizontalResolution
    VerticalResolution
    IsAnimated
    ActiveFrame
    FrameCount
    FileExtension
    HasTransparency
    PixelDepth
End Enum

Function GetImageProperties(ByVal File As Variant, ByVal PropertyType As ImagePropertyTypeEnum)
 
    Dim TargetImage As Object
 
    If TypeName(File) = "IImageFile" Then
        Set TargetImage = File
    ElseIf TypeName(File) = "String" Then
        If Len(Dir(File)) = 0 Then Exit Function
        ' Late-binding option
        Set TargetImage = CreateObject("WIA.ImageFile")
        ' Early-binding option
        ' Dim TargetImage As New WIA.ImageFile
        TargetImage.LoadFile File
    End If
 
    With TargetImage
        GetImageProperties = Trim(Array(.Width, .Height, .HorizontalResolution, .VerticalResolution, _
                                        CBool(.IsAnimated), .ActiveFrame, .FrameCount, .FileExtension, _
                                        CBool(.IsAlphaPixelFormat), .PixelDepth)(PropertyType))
    End With
 
    Set TargetImage = Nothing
 
End Function

Sub BasicImageProperties_Test(ByVal File As Variant)
 
    Dim PropertyHeadings  As Variant
    PropertyHeadings = Array("Width (pixels)", "Height (pixels)", "Horizontal Resolution", "Vertical Resolution", _
                             "Animated", "Active Frame", "Frame Count", "File Extension", _
                             "Is Alpha Pixel Format", "Pixel Depth")
    Dim Counter As Long
 
    Dim Heading As String * 22
    For Counter = LBound(PropertyHeadings) To UBound(PropertyHeadings)
        Heading = PropertyHeadings(Counter)
        Debug.Print Counter + 1, Heading, GetImageProperties(File, Counter)
    Next
 
End Sub

Function GetImageFileFromURL(ByVal TargetURL As String) As Object
    Dim HTTP            As Object
 
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", TargetURL, False
    HTTP.send

    If HTTP.Status = 200 Then
        With CreateObject("WIA.Vector")
            .BinaryData = HTTP.responsebody
            Set GetImageFileFromURL = .ImageFile
        End With
    End If

    Set HTTP = Nothing
End Function


Attribute VB_Name = "modWIA_ImageCreation"
                                                                                                                                                                                        ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||      modWIA_ImageCreation (v1.0)      ||||||||||||||||||||||||||||||||||                                                                                     ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                        ' _
AUTHOR:   Kallun Willock
                                                                                                                                                                                        ' _
URL:      https://www.mrexcel.com/board/threads/right-mouse-click-menu-extra-functionality.1223424/                                                                                     ' _
          https://www.reddit.com/r/vba/comments/xo8zs5/convert_multipage_pdf_to_multipage_tiff/
                                                                                                                                                                                        ' _
PURPOSE:  A simple class module that leverages the functionality of the WIA COM Object for image manipulation/conversion.                                                               ' _
          It was developed to respond to the specific requirements set out in the thread at the above-referenced URL.
                                                                                                                                                                                        ' _
LICENSE:  MIT                                                                                                                                                                           ' _
VERSION:  0.1        26/09/2022         Published TIFF routine on Reddit                                                                                                                ' _
          0.2        30/11/2022         Published CreateFaceIDBMP routine on MrExcel.com                                                                                                ' _
          1.0        26/05/2023         Published on Github                                                                                                                             ' _
TODO:     [ ] Error Handling

Option Explicit

' https://www.mrexcel.com/board/threads/right-mouse-click-menu-extra-functionality.1223424/

Function CreateFaceIDBMP(ByVal TargetColor As Long) As StdPicture
    Dim ImgVector As Object
    Dim Counter As Long
    Dim Node As Object
    
    Set ImgVector = CreateObject("WIA.Vector")
    Set Node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    
    Node.DataType = "bin.base64"
    Node.Text = "Qk02AwAAAAAAADYAAAAoAAAAEAAAABAAAAABABgAAAAAAAADAAAAAAAAAAAAAAAAAAAAAAAA" _
    & WorksheetFunction.Rept("yMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjI" & vbNewLine, 14) _
    & "yMjIyMjIyMjIyMjI"
    
    With ImgVector
        .BinaryData = Node.nodeTypedValue
        For Counter = 55 To UBound(.BinaryData) Step 3
            .Item(Counter) = (TargetColor \ 65536) Mod 256
            .Item(Counter + 1) = (TargetColor \ 256) Mod 256
            .Item(Counter + 2) = (TargetColor Mod 256)
        Next
        Set CreateFaceIDBMP = .Picture
    End With
    Set Node = Nothing
End Function


' https://www.reddit.com/r/vba/comments/xo8zs5/convert_multipage_pdf_to_multipage_tiff/

Sub CreateTIFF()
    
    Dim Page1   As Object
    Dim Page2   As Object
    Dim Page3   As Object
    Dim IMG     As Object
    
    Set Page1 = CreateObject("WIA.ImageFile")
    Set Page2 = CreateObject("WIA.ImageFile")
    Set Page3 = CreateObject("WIA.ImageFile")
    Set IMG = CreateObject("WIA.ImageProcess")
    
    Page1.LoadFile "D:\page1.png"
    Page2.LoadFile "D:\page2.png"
    Page3.LoadFile "D:\page3.png"
    
    ' Add a page/frame filter - Page 2
    IMG.Filters.Add IMG.FilterInfos("Frame").FilterID
    Set IMG.Filters(IMG.Filters.Count).Properties("ImageFile") = Page2
    
    ' Add a page/frame filter - Page 3
    IMG.Filters.Add IMG.FilterInfos("Frame").FilterID
    Set IMG.Filters(IMG.Filters.Count).Properties("ImageFile") = Page3
    
    ' Add a converter filter for the TIFF file format
    IMG.Filters.Add IMG.FilterInfos("Convert").FilterID
    IMG.Filters(IMG.Filters.Count).Properties("FormatID") = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
    
    ' Apply the filters to the first image / page 1
    Set Page1 = IMG.Apply(Page1)
    
    ' Save Page 1 as a TIFF file
    Page1.SaveFile "D:\IDislikeTIF.tif"
    
    Set Page1 = Nothing
    Set Page2 = Nothing
    Set Page3 = Nothing
    Set IMG = Nothing

End Sub



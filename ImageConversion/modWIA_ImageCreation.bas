Attribute VB_Name = "modWIA_ImageCreation"
                                                                                                                                                                                        ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||      modWIA_ImageCreation (v0.1)      ||||||||||||||||||||||||||||||||||                                                                                     ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                        ' _
AUTHOR:   Kallun Willock                                                                                                                                                                ' _
URL:      https://www.mrexcel.com/board/threads/right-mouse-click-menu-extra-functionality.1223424/                                                                                     ' _
PURPOSE:  A simple class module that leverages the functionality of the WIA COM Object for image manipulation/conversion.                                                               ' _
          It was developed to respond to the specific requirements set out in the thread at the above-referenced URL.                                                                   ' _
LICENSE:  MIT                                                                                                                                                                           ' _
VERSION:  1.0        08/04/2023         Published v1.0 on Mr.Excel forum.                                                                                                               ' _
          1.1        18/05/2023         Added comments for publication Github                                                                                                           ' _
TODO:     [ ] Error Handling                                                                                                                                                            ' _

        
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



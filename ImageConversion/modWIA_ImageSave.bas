Attribute VB_Name = "modWIA_ImageSave"
                                                                                                                                                                                        ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||         modWIA_ImageSave (v1.0)       ||||||||||||||||||||||||||||||||||                                                                                     ' _
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
#If VBA7 Then
    Private Declare PtrSafe Function DispCallFunc Lib "oleAut32.dll" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
#Else
    Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal oVft As Long, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
#End If

' https://www.vbforums.com/showthread.php?889248-VB6-Convert-a-picture-to-PNG-byte-array-in-memory

Public Function SaveAsPng(pPic As stdole.IPicture) As Byte()
    Const adTypeBinary  As Long = 1
    Const wiaFormatPNG  As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Const CC_STDCALL    As Long = 4
    Dim oStream         As Object ' ADODB.Stream
    Dim oImageFile      As Object ' WIA.ImageFile
    Dim IID_IStream(3)  As Long
    Dim pStream         As IUnknown
    Dim vParams(0 To 1) As Variant
    Dim vType(0 To 1)   As Integer
    Dim vPtr(0 To 1)    As LongPtr
    
    '--- load pPic in WIA.ImageFile
    Do While oImageFile Is Nothing
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Type = adTypeBinary
        oStream.Open
        '--- call IUnknown::QI on oStream for IStream interface and store in pStream
        IID_IStream(0) = &HC
        IID_IStream(2) = &HC0
        IID_IStream(3) = &H46000000
        vParams(0) = VarPtr(IID_IStream(0))
        vParams(1) = VarPtr(pStream)
        vType(0) = VarType(vParams(0))
        vType(1) = VarType(vParams(1))
        vPtr(0) = VarPtr(vParams(0))
        vPtr(1) = VarPtr(vParams(1))
        Call DispCallFunc(ObjPtr(oStream), 0, CC_STDCALL, vbLong, UBound(vParams) + 1, vType(0), vPtr(0), Empty)
        '--- NO magic anymore, only business as usual
        pPic.SaveAsFile ByVal ObjPtr(pStream), True, 0
        If oStream.Size = 0 Then
            GoTo QH
        End If
        oStream.Position = 0
        With CreateObject("WIA.Vector")
            .BinaryData = oStream.Read
            If pPic.Type <> 1 Then
                '--- this converts pPic to vbPicTypeBitmap subtype
                Set pPic = .Picture
            Else
                Set oImageFile = .IMageFile
            End If
        End With
    Loop
    '--- serialize WIA.ImageFile to PNG file format
    With CreateObject("WIA.ImageProcess")
        .Filters.Add .FilterInfos("Convert").FilterID
        .Filters(.Filters.Count).Properties("FormatID").Value = wiaFormatPNG
        SaveAsPng = .Apply(oImageFile).FileData.BinaryData
    End With
QH:
End Function
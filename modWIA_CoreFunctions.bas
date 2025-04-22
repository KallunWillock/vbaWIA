


Function GetFramePicture(ByRef ImageFile as WIA.ImageFile, ByVal FrameIndex as Long) as IPictureDisp
    Dim IPic As IPictureDisp
    with ImageFile
        .ActiveFrame = CLng(FrameIndex)
        set IPic = .ARGBData.ImageFile(.Width, .Height).FileData.Picture
    End With
    Set GetFramePicture = IPic
End Function

Function LoadPictureEx(ByVal Filename as string) as IPictureDisp
    with CreateObject("WIA.ImageFile")
        .LoadFile(Filename)
        Set LoadPictureEx = .FileData.Picture
    End With
End Function
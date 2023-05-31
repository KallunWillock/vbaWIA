'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     
'||||||||||||||||||||||||||           ConvertToPNG.vbs            ||||||||||||||||||||||||||||||||||                                                                                     
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

'AUTHOR:   Kallun Willock                                                                                             

'PURPOSE:  A demonstration of WIA using VBScript. This script converts image files dragged and dropped
'          onto the VBScript file into the PNG Image Format.

'LICENSE:  MIT                                                                                                        

Sub Main()
    Const wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
    Dim Img                 ' As ImageFile
    Dim objArgs             ' Handles the drag/dropped files
    Dim I
    Dim Filename
    
    ' WScript.Arguments comprises detail of any arguments passed
    ' to the script at the command line, or the file paths/names
    ' of any files dropped onto the VBS file in the Windows Explorer.
    
    Set objArgs = WScript.Arguments
    Set Img = CreateObject("WIA.ImageFile")
    
    For I = 0 To objArgs.Count - 1
        Filename = objArgs.Item(I)
        Extension = GetExtension(Filename)
        If InStr("jpeg|jpg|gif|bmp|dib|tiff|png", Extension) Then
            Img.LoadFile Filename
            If Img.FormatID <> wiaFormatPNG Then
                Dim IP      ' As ImageProcess
                Set IP = CreateObject("Wia.ImageProcess")
                IP.Filters.Add IP.FilterInfos("Convert").FilterID
                IP.Filters(1).Properties("FormatID").Value = wiaFormatPNG
                Set Img = IP.Apply(Img)
                Img.SaveFile ReplaceExtension(Filename, "png")
            End If
         End If
     Next
End Sub
 
Function ReplaceExtension(Filename, NewExtension)
    Dim Parts
    Parts = Split(Filename, ".")
    Parts(UBound(Parts)) = NewExtension
    ReplaceExtension = Join(Parts, ".")
End Function
 
Function GetExtension(Filename)
    Dim Parts
    Parts = Split(Filename, ".")
    GetExtension = LCase(Parts(UBound(Parts)))
End Function
 


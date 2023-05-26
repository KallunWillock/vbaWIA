Attribute VB_Name = "modWIA_ImageLoad"

                                                                                                                                                                                        ' _
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
||||||||||||||||||||||||||         modWIA_ImageLoad (v1.0)       ||||||||||||||||||||||||||||||||||                                                                                     ' _
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

' You could even rename the function as LoadPicture, and force VBA to use this custom function as the go-to
' routine rather than the inbuilt (limited) routine without breaking existing code. The other benefit is that
' it retains alpha channel transparency, depending on the control you're loading it into.
       
Function LoadImage(ByVal Filename As String) As StdPicture
        With CreateObject("WIA.ImageFile")
                .LoadFile Filename
                Set LoadImage = .FileData.Picture
        End With
End Function

Function ByteArrayToStdPicture(ByVal ImageData As Variant) As StdPicture
    
    With CreateObject("WIA.Vector")
        .BinaryData = ImageData
        Set ByteArrayToStdPicture = .Picture
    End With
    
End Function

Function Base64toStdPicture(ByVal Base64Code As String) As StdPicture
        
    Dim Node As Object
    Set Node = CreateObject("Msxml2.DOMDocument.3.0").createElement("base64")
    
    Node.DataType = "bin.base64"
    Node.Text = Base64Code
    
    With CreateObject("WIA.Vector")
        .BinaryData = Node.NodeTypedValue
        Set Base64toStdPicture = .Picture
    End With
    
    Set Node = Nothing
    
End Function

Function GetImageFromURL(ByVal TargetURL As String) As StdPicture

    Dim HTTP            As Object
    
    Set HTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    HTTP.Open "GET", TargetURL, False
    HTTP.send

    If HTTP.Status = 200 Then
        With CreateObject("WIA.Vector")
            .BinaryData = HTTP.responseBody
            Set GetImageFromURL = .Picture
        End If
    End If

    Set HTTP = Nothing
    
End Function


'Adding PNG images to Userforms/Userform Controls in design-mode
'It is possible to add PNG images to the picture property of a UserForm/UserForm Control at design time by leveraging the Microsoft Visual Basic for Application Extensibility Library - you will need to add a reference to it by adding it through Tools -> References -> select the Microsoft Visual Basic for Application Extensibility Library 5.x library. The subroutine below is an example of code you can execute in design-mode to load PNG images into userforms/userform controls.

'Assuming that you want to load a PNG file (located at "D:\sample.png") into a label control (called Label1) on userform (called UserForm1), execute the following command in the Immediate Window:

'VBA Code:
'PNG2Picture "D:\Sample.png", "Label1", "UserForm1"

'If you leave the filename argument empty, the routine will display an Open File dialogbox from which the image file can be selected:

'VBA Code:
'PNG2Picture , "Label1", "UserForm1"

'If you leave the controlname argument empty, the routine will create a new label control in the relevent userform with the name "NewImageControl":
'
VBA Code:
'PNG2Picture "D:\Sample.png", , "UserForm1"

'If you leave the UserForm argument empty, the routine will assume that you have selected the preferred UserForm. Accordingly, if you leave all three arguments empty, you will be prompted for an image file, and a new control with the name "NewImageControl" will be created in the selected userform:

'   VBA Code:
'   PNG2Picture
Sub PNG2Picture(Optional ByVal Filename As String, Optional ByVal ControlName As String, Optional ByVal VBCName As Variant)
    
    If Len(Filename) = 0 Then
        Filename = Application.GetOpenFilename(FileFilter:="Image Files (*.PNG; *.BMP; *.DIB; *.GIF; *.JPG; *.JPEG), *.PNG;*.BMP;*.DIB;*.GIF;*.JPG;*.JPEG,PNG Image (*.PNG),*.png,Bitmaps Image (*.BMP; *.DIB), *.bmp;*.dib,JPEG Image (*.JPG; *.JPEG), *.jpg;*.jpeg,GIF Image (*.GIF), *.gif, All files (*.*), *.*")
        If Filename = "False" Then Exit Sub
    End If
    
    Dim TargetObject    As MSForms.Control
    Dim VBP             As VBProject
    Dim VBC             As VBComponent
    Dim DesignTool      As Object

    Set VBP = Application.VBE.ActiveVBProject
    If IsMissing(VBCName) Then
        Set VBC = Application.VBE.SelectedVBComponent
    Else
        Set VBC = Application.VBE.ActiveVBProject.VBComponents(VBCName)
    End If
    
    If VBC.HasOpenDesigner Then
        Set DesignTool = VBC.Designer
        On Error Resume Next
        If Len(ControlName) Then
            Set TargetObject = DesignTool.Controls(ControlName)
        Else
            Set TargetObject = DesignTool.Controls.Add("Forms.Label.1", "NewImageControl")
        End If
        On Error GoTo 0
        If Not TargetObject Is Nothing Then
            With TargetObject
                .Caption = " "
                .BackStyle = 0
                .BorderStyle = 0
                .AutoSize = True
            End With
            With CreateObject("WIA.ImageFile")
                .LoadFile Filename
                TargetObject.Picture = .FileData.Picture
            End With
        End If
    End If
End Sub

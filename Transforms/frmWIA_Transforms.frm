' In a Userform, with four command buttons (btnLoadPic, btnRotate90, btnScale, btnSave) and one imagebox (Image1):
Option Explicit

Private WIAExample As New clsWIA

Private Sub btnLoadPic_Click()
    Dim Filename As Variant
    Filename = Application.GetOpenFilename(Title:="Add Employee Image")
    If Filename = False Then Exit Sub
    WIAExample.Source = Filename
    Me.Image1.Picture = WIAExample.Picture
End Sub

Private Sub btnRotate90_Click()
    WIAExample.RotateImage RotationEnum.Rotation90
    Me.Image1.Picture = WIAExample.Picture
End Sub

Private Sub btnSave_Click()
    WIAExample.SaveImage PNG
End Sub

Private Sub btnScale_Click()
    WIAExample.ScaleImage 100, 100
    Me.Image1.Picture = WIAExample.Picture
End Sub
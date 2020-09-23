VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSkinForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim SkinedForm As Form

Dim BackgroundPic As IPictureDisp
Dim UpLeftPic As IPictureDisp
Dim UpRightPic  As IPictureDisp
Dim UpMiddlePic As IPictureDisp
Dim LeftSidePic As IPictureDisp
Dim RightSidePic As IPictureDisp
Dim BottomLeftPic As IPictureDisp
Dim BottomMiddlePic As IPictureDisp
Dim BottomRightPic As IPictureDisp

Dim ButtonUpPic As IPictureDisp
Dim ButtonOverPic  As IPictureDisp
Dim ButtonDownPic As IPictureDisp

Dim MaxUpPic As IPictureDisp
Dim MaxOverPic As IPictureDisp
Dim MaxDownPic As IPictureDisp

Dim MinUpPic As IPictureDisp
Dim MinOverPic As IPictureDisp
Dim MinDownPic As IPictureDisp

Dim CloseUpPic As IPictureDisp
Dim CloseOverPic As IPictureDisp
Dim CloseDownPic As IPictureDisp

Dim Background As Image
Dim UpLeft As Image
Dim BackLeft As Image
Dim LeftSide As Image
Dim UpMiddle As Image
Dim RightSide As Image
Dim UpRight As Image
Dim BottomRight As Image
Dim BottomLeft As Image
Dim BottomMiddle As Image


Sub StretchIt()

    Background.Stretch = True

    UpLeft.Stretch = True
    UpRight.Stretch = True
    UpMiddle.Stretch = True

    LeftSide.Stretch = True
    RightSide.Stretch = True

    BottomLeft.Stretch = True
    BottomMiddle.Stretch = True
    BottomRight.Stretch = True

End Sub

Sub UnStretchIt()

    Background.Stretch = False

    UpLeft.Stretch = False
    UpRight.Stretch = False
    UpMiddle.Stretch = False

    LeftSide.Stretch = False
    RightSide.Stretch = False

    BottomLeft.Stretch = False
    BottomMiddle.Stretch = False
    BottomRight.Stretch = False

End Sub

Sub GFX()
On Error Resume Next
    Background.Move 0, 0, SkinedForm.Width, SkinedForm.Height
    
    UpLeft.Move 0, 0, UpLeft.Width, UpLeft.Height
    UpMiddle.Move UpLeft.Width, 0, SkinedForm.Width - UpLeft.Width - UpRight.Width - 90, UpMiddle.Height
    UpRight.Move SkinedForm.Width - UpRight.Width - 90, 0, UpRight.Width, UpRight.Height
    
    LeftSide.Move 0, UpLeft.Height, LeftSide.Width, SkinedForm.Height - UpLeft.Height - BottomLeft.Height
    RightSide.Move SkinedForm.Width - RightSide.Width - 90, UpRight.Height, RightSide.Width, SkinedForm.Height - UpRight.Height - BottomRight.Height
    BottomLeft.Move 0, SkinedForm.Height - BottomLeft.Height, BottomLeft.Width, BottomLeft.Height
    
    BottomMiddle.Move BottomLeft.Width, SkinedForm.Height - BottomMiddle.Height - 90, SkinedForm.Width - BottomLeft.Width - BottomRight.Width - 90, BottomMiddle.Height
    BottomLeft.Move 0, SkinedForm.Height - BottomLeft.Height - 90, BottomLeft.Width, BottomLeft.Height
    BottomRight.Move SkinedForm.Width - BottomRight.Width - 90, SkinedForm.Height - BottomRight.Height - 90, BottomRight.Width, BottomRight.Height
       
End Sub

Sub CreateImages()

Set Background = SkinedForm.Controls.Add("VB.Image", "Background", SkinedForm)
Set UpLeft = SkinedForm.Controls.Add("VB.Image", "UpLeft", SkinedForm)
Set UpMiddle = SkinedForm.Controls.Add("VB.Image", "UpMiddle", SkinedForm)
Set UpRight = SkinedForm.Controls.Add("VB.Image", "UpRight", SkinedForm)
Set LeftSide = SkinedForm.Controls.Add("VB.Image", "LeftSide", SkinedForm)
Set RightSide = SkinedForm.Controls.Add("VB.Image", "RightSide", SkinedForm)
Set BottomLeft = SkinedForm.Controls.Add("VB.Image", "BottomLeft", SkinedForm)
Set BottomRight = SkinedForm.Controls.Add("VB.Image", "BottomRight", SkinedForm)
Set BottomMiddle = SkinedForm.Controls.Add("VB.Image", "BottomMiddle", SkinedForm)

Background.Visible = True
UpLeft.Visible = True
UpMiddle.Visible = True
UpRight.Visible = True
LeftSide.Visible = True
RightSide.Visible = True
BottomLeft.Visible = True
BottomRight.Visible = True
BottomMiddle.Visible = True

End Sub

Sub SkinForm(Form As Form, Settings As String)

Set SkinedForm = Form
CreateImages
UnStretchIt
LoadImages Settings
SetPictures
StretchIt
GFX
MakeResizeable
MoveForm
End Sub

Sub LoadImages(Settings As String)
Dim Tag As String, Tmp As String
Dim Parameter As String, Setting As String
Dim Path As String
Dim F As Long

On Error GoTo Error

Path = Mid(Settings, 1, InStrRev(Settings, "\"))
F = FreeFile

Open Settings For Input As #F

Do Until EOF(F)
Line Input #F, Tmp
If Len(Tmp) <> 0 Then
If Mid(Tmp, 1, 1) = "[" Then
    Tag = Tmp
Else:
    Parameter = Split(Tmp, "=")(0)
    Setting = Split(Tmp, "=")(1)

Select Case LCase(Tag)
    Case "[picture]"

    Select Case LCase(Parameter)
        Case "background"
            Set BackgroundPic = LoadPicture(Path & Setting)
        Case "upleft"
            Set UpLeftPic = LoadPicture(Path & Setting)
        Case "upright"
            Set UpRightPic = LoadPicture(Path & Setting)
        Case "upmiddle"
            Set UpMiddlePic = LoadPicture(Path & Setting)
        Case "leftside"
            Set LeftSidePic = LoadPicture(Path & Setting)
        Case "rightside"
            Set RightSidePic = LoadPicture(Path & Setting)
        Case "bottomleft"
            Set BottomLeftPic = LoadPicture(Path & Setting)
        Case "bottomright"
            Set BottomRightPic = LoadPicture(Path & Setting)
        Case "bottommiddle"
            Set BottomMiddlePic = LoadPicture(Path & Setting)
    End Select
    
End Select
    

End If
End If
Loop
Close #F

Exit Sub
Error:
MsgBox "Error Loading: '" & Setting & "'", vbApplicationModal + vbCritical, "Error"
End Sub

Sub SetPictures()
Background.Picture = BackgroundPic
UpLeft.Picture = UpLeftPic
UpMiddle.Picture = UpMiddlePic
UpRight.Picture = UpRightPic
LeftSide.Picture = LeftSidePic
RightSide.Picture = RightSidePic
BottomLeft.Picture = BottomLeftPic
BottomMiddle.Picture = BottomMiddlePic
BottomRight.Picture = BottomRightPic

Background.ZOrder 1
End Sub

Sub MoveForm()
Do
If GetActiveWindow() = SkinedForm.hwnd And GetAsyncKeyState(MOUSEEVENTF_LEFTDOWN) <> 0 Then
ReleaseCapture
SendMessage SkinedForm.hwnd, &HA1, 2, 0&
End If
DoEvents
Loop
End Sub

Sub MakeResizeable()
Call SetWindowLong(SkinedForm.hwnd, GWL_STYLE, GetWindowLong(SkinedForm.hwnd, GWL_STYLE) + WS_DLGFRAME)
End Sub



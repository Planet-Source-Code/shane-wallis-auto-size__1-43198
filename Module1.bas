Attribute VB_Name = "Module1"
Option Explicit


'   Resolution Resize and Run Time Control Resize.
'Module for automatically resizing forms and
'controls with varying screen resolutions. This is
'an adaptation of a Microsoft Knowledge Base
'article. The example worked as it was but half the
'code was on the form. I wanted a complete module
'to add to any app easily. To use it simply add
'"Call AdjustForm(Me)" to the Form_Load event
'and "Call FormResize(Me)" to the Form_Resize
'event.  Also change the design time resolution
'values. It is coded to 640x480 since my video
'adapter will not support higher at 16 bit color.
'The Microsoft article said it was for VB5/6 but
'I have only VB4. If you have trouble, make one for
'yourself...use at your own risk, else e-mail me
'at nwsr2@netscape.net.        No API's

Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

Type FRMSIZE
   Height As Long
   Width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer
Dim ScaleFactorX As Single, ScaleFactorY As Single

 
Sub Resize_For_Resolution(ByVal SFX As Single, ByVal SFY As Single, MyForm As Form)
Dim I As Integer
Dim SFFont As Single
SFFont = (SFX + SFY) / 2
On Error Resume Next
With MyForm
  For I = 0 To .Count - 1
   If TypeOf .Controls(I) Is ComboBox Then
     .Controls(I).Left = .Controls(I).Left * SFX
     .Controls(I).Top = .Controls(I).Top * SFY
     .Controls(I).Width = .Controls(I).Width * SFX
   Else
     .Controls(I).Move .Controls(I).Left * SFX, _
      .Controls(I).Top * SFY, _
      .Controls(I).Width * SFX, _
      .Controls(I).Height * SFY
   End If
     .Controls(I).FontSize = .Controls(I).FontSize * SFFont
  Next I
  If RePosForm Then
     .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
  End If
End With
End Sub


Public Sub FormResize(TheForm As Form)
Dim ScaleFactorX As Single, ScaleFactorY As Single
If Not DoResize Then
   DoResize = True
   Exit Sub
End If
RePosForm = False
ScaleFactorX = TheForm.Width / MyForm.Width
ScaleFactorY = TheForm.Height / MyForm.Height
Resize_For_Resolution ScaleFactorX, ScaleFactorY, TheForm
MyForm.Height = TheForm.Height
MyForm.Width = TheForm.Width
End Sub

Public Sub AdjustForm(TheForm As Form)
Dim Res As String ' Returns resolution of system
' Put the design time resolution in here
DesignX = 640
DesignY = 480
RePosForm = True
DoResize = False
Xtwips = Screen.TwipsPerPixelX
Ytwips = Screen.TwipsPerPixelY
Ypixels = Screen.Height / Ytwips
Xpixels = Screen.Width / Xtwips
ScaleFactorX = (Xpixels / DesignX)
ScaleFactorY = (Ypixels / DesignY)
TheForm.ScaleMode = 1
Resize_For_Resolution ScaleFactorX, ScaleFactorY, TheForm
Res = Str$(Xpixels) + "  by " + Str$(Ypixels)
Debug.Print Res
MyForm.Height = TheForm.Height
MyForm.Width = TheForm.Width
 
End Sub





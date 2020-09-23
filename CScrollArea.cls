VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScrollArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents VBar As VScrollBar
Attribute VBar.VB_VarHelpID = -1
Public WithEvents HBar As HScrollBar
Attribute HBar.VB_VarHelpID = -1
Public WithEvents InnerPicture As PictureBox
Attribute InnerPicture.VB_VarHelpID = -1
Public WithEvents FramePicture As PictureBox
Attribute FramePicture.VB_VarHelpID = -1

Const DFC_SCROLL = 3                    ' Scroll bar

Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Sub SetScrollBar()
    ' See if we need the vertical scroll bar.
    If InnerPicture.Height <= FramePicture.ScaleHeight Then
        VBar.Enabled = False
    Else
        ' Set scroll bar properties.
        InnerPicture.top = 0
        VBar.Min = 0
        VBar.Max = FramePicture.ScaleHeight - InnerPicture.Height
        VBar.LargeChange = FramePicture.ScaleHeight
        VBar.SmallChange = FramePicture.ScaleHeight / 5
        VBar.Enabled = True
    End If

    ' See if we need the vertical scroll bar.
    If InnerPicture.Width <= FramePicture.ScaleWidth Then
        HBar.Enabled = False
    Else
        ' Set scroll bar properties.
        InnerPicture.Left = 0
        HBar.Min = 0
        HBar.Max = FramePicture.ScaleWidth - InnerPicture.Width
        HBar.LargeChange = FramePicture.ScaleWidth
        HBar.SmallChange = FramePicture.ScaleWidth / 5
        HBar.Enabled = True
    End If
End Sub
Private Sub VBar_Change()
    InnerPicture.top = VBar.Value
End Sub
Private Sub VBar_Scroll()
    InnerPicture.top = VBar.Value
End Sub
Private Sub HBar_Change()
    InnerPicture.Left = HBar.Value
End Sub
Private Sub HBar_Scroll()
    InnerPicture.Left = HBar.Value
End Sub
Sub ReSizeArea()
    VBar.TabStop = False: HBar.TabStop = False
    InnerPicture.Left = ((FramePicture.Width) - InnerPicture.Width) / 2
    InnerPicture.top = (FramePicture.Height - InnerPicture.Height) / 2
    SetScrollBar
End Sub


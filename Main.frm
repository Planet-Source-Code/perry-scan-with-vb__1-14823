VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan"
   ClientHeight    =   6768
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11424
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6768
   ScaleWidth      =   11424
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   5892
      Left            =   120
      ScaleHeight     =   5844
      ScaleWidth      =   11244
      TabIndex        =   2
      Top             =   840
      Width           =   11292
      Begin VB.HScrollBar HScroll 
         Height          =   252
         Left            =   0
         TabIndex        =   4
         Top             =   5640
         Width           =   11052
      End
      Begin VB.VScrollBar VScroll 
         Height          =   5652
         LargeChange     =   5
         Left            =   11040
         TabIndex        =   3
         Top             =   0
         Width           =   252
      End
      Begin VB.PictureBox picScan 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1212
         Left            =   0
         ScaleHeight     =   1212
         ScaleWidth      =   1692
         TabIndex        =   5
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan!"
      Height          =   612
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   5532
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select TWAIN Source"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5532
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The TWAIN32d.dll should be in the System directory
Private Declare Function TWAIN_AcquireToFilename Lib "TWAIN32d.DLL" (ByVal hwndApp As Long, ByVal bmpFileName As String) As Integer
Private Declare Function TWAIN_IsAvailable Lib "TWAIN32d.DLL" () As Long
Private Declare Function TWAIN_SelectImageSource Lib "TWAIN32d.DLL" (ByVal hwndApp As Long) As Long
Dim ScrollAreaScan As CScrollArea

Private Sub cmdScan_Click()
Dim Ret As Long, PictureFile As String
PictureFile = App.Path & "\temp.bmp"
'PicturFile is the temporary file "temp.bmp"
'In "temp.bmp" the image will stored until the end of the action
Ret = TWAIN_AcquireToFilename(Me.hwnd, PictureFile)
If Ret = 0 Then
'If the scan is successful
picScan.Picture = LoadPicture(PictureFile)
'Load the temporary picture file
ScrollAreaScan.ReSizeArea
'Resize the picture control
Kill PictureFile
'Delete the temporary picture file
Else
MsgBox "Scan not successful!", vbCritical, "Scanning"
End If
End Sub

Private Sub cmdSelect_Click()
TWAIN_SelectImageSource (Me.hwnd)
End Sub

Private Sub Form_Load()
Set ScrollAreaScan = New CScrollArea
Set ScrollAreaScan.VBar = VScroll
Set ScrollAreaScan.HBar = HScroll
Set ScrollAreaScan.InnerPicture = picScan
Set ScrollAreaScan.FramePicture = Picture1
End Sub


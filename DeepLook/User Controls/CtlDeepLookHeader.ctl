VERSION 5.00
Begin VB.UserControl ucDeepLookHeader 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   Picture         =   "CtlDeepLookHeader.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   5355
   ToolboxBitmap   =   "CtlDeepLookHeader.ctx":08A5
End
Attribute VB_Name = "ucDeepLookHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  .======================================.
' /         DeepLook Project Scanner       \
' |       By Dean Camera, 2003 - 2005      |
' \  Visual Basic Project Scanning Engine  /
'  '======================================'
' / Most of this project is now commented  \
' \           to help developers.          /
'  '======================================'

' Used to minimize memory requirements for DeepLook by only storing one logo

Private Sub UserControl_Resize()
    On Error Resume Next
    
    UserControl.Height = 370
    UserControl.Width = UserControl.Parent.Width
End Sub

Sub ResizeMe()
    UserControl_Resize
End Sub

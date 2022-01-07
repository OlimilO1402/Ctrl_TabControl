Attribute VB_Name = "MApp"
Option Explicit
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Sub Main()
    FMain.Show
End Sub


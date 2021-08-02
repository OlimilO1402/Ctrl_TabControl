Attribute VB_Name = "MApp"
Option Explicit
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Sub Main()
    Form1.Show
End Sub


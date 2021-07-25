Attribute VB_Name = "ModConstructors"
Option Explicit

Public Function New_TabControl(myOwner As Form, MyContainer As PictureBox, MyName As String) As TabControl
  Set New_TabControl = New TabControl
  Call New_TabControl.NewC(myOwner, MyContainer, MyName)
End Function

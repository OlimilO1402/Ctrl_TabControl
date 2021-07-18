Attribute VB_Name = "MNew"
Option Explicit

'##############################'    TabControl   '##############################'
Public Function TabControl(MyOwner As Form, MyContainer As PictureBox, MyName As String) As TabControl
    Set TabControl = New TabControl
    Call TabControl.New_(MyOwner, MyContainer, MyName)
End Function
'Public Function New_TabPage(Name As String, Optional Ctrl As PictureBox = Nothing) As TabPage
'  Set New_TabPage = New TabPage
'  New_TabPage.Text = Name
'  If Not Ctrl Is Nothing Then
'    Ctrl.BorderStyle = 0
'    Call New_TabPage.Controls.Add(Ctrl)
'  End If
'  'call New_TabPage.
'End Function


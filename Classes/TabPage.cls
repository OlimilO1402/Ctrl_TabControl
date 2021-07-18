VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Public Class TabPage
'          Inherits System.Windows.Forms.Panel
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt eine einzelne Tabseite in System.Windows.Forms.TabControl dar.
Option Explicit
Private Const TCM_FIRST             As Long = &H1300  '{ Tab control messages }   Dez: 4864
Private Const TCM_GETITEMA          As Long = (TCM_FIRST + 5)
Private Const TCM_SETITEMA          As Long = (TCM_FIRST + 6)
Private Const TCM_INSERTITEMA       As Long = (TCM_FIRST + 7)
Private Const TCM_DELETEITEM        As Long = (TCM_FIRST + 8)
Private Const TCM_GETITEM           As Long = TCM_GETITEMA
Private Const TCM_SETITEM           As Long = TCM_SETITEMA
Private Const TCM_INSERTITEM        As Long = TCM_INSERTITEMA
Private Type TCITEM 'TabControlItem = Tabber
    mask        As Long
    dwState     As Long
    dwStateMask As Long
    pszText     As String
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long
End Type

Private mTabStripName As String
Private mName         As String
Private mText         As String
Private mParent       As Form
Private mPicBPage     As PictureBox  'dies ist die Seite, sie dient als Container für alles Mögliche
Private mVisible      As Boolean
Private mImageIndex   As Long
Private mToolTipText  As String
Private mDebugLabel   As Label
Private WithEvents mContainer As PictureBox 'Container für alle Seiten
Attribute mContainer.VB_VarHelpID = -1
Private WithEvents mControls  As ControlCollection
Attribute mControls.VB_VarHelpID = -1
Private mSTPX As Long, mSTPY As Long
'Public Dock As New DockStyle
Public mTCHandle      As Long 'Handle des TabControls
Public mIndex         As Long 'der eigene Index und der des Tabs - beginnt bei 0

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetParent Lib "user32.dll" ( _
    ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Class_Initialize()
    Set mControls = New ControlCollection
    mTCHandle = 0
End Sub
Private Sub Class_Terminate()
    Set mPicBPage = Nothing
    Set mParent = Nothing
    Set mContainer = Nothing
End Sub

' #################### '    Public Procedures    ' #################### '
Public Sub Create(MyOwner As Object, MyContainer As Object, MyTabStripName As String, MyText As String)
    mText = MyText
    mTabStripName = MyTabStripName
    mName = mTabStripName & "mPicBPage" & ReplNonAlpNum(mText)
    Set mParent = MyOwner
    Set mContainer = MyContainer
    'mContainer.BorderStyle = 0 'halt hier wird bereits ein Resize ausgelöst
    'mContainer.Appearance = 1
    If mPicBPage Is Nothing Then Set mPicBPage = mParent.Controls.Add("VB.PictureBox", mName, mContainer)
    mPicBPage.ScaleMode = mContainer.ScaleMode
    mPicBPage.BorderStyle = 0
    mPicBPage.Appearance = 1
    If mContainer.ScaleMode = vbTwips Then
        mSTPX = Screen.TwipsPerPixelX: mSTPY = Screen.TwipsPerPixelY
    Else
        mSTPX = 1: mSTPY = 1
    End If
    'mPicBPage.BackColor = &HFEFCFC
    'Call ResizeMe
    mPicBPage.Visible = True
    'If mDebugLabel Is Nothing Then Set mDebugLabel = mParent.Controls.Add("VB.Label", mName & "Lbl", mPicBPage)
    If Not mDebugLabel Is Nothing Then
        mDebugLabel.AutoSize = True
        mDebugLabel.Caption = mName + mText
        mDebugLabel.Visible = True
    End If
    mContainer.BorderStyle = 0 'halt hier wird bereits ein Resize ausgelöst
    mContainer.Appearance = 1
End Sub
Private Function ReplNonAlpNum(Text As String) As String
    Dim i  As Long
    Dim ch As String * 1
    Dim aa As Long
    'Alle Sonderzeichen im Namen entfernen und durch "_" ersetzen
    'alle "\&%$§!"
    'MsgBox "Asc 0: " & CStr(Asc("0")) & " " & "Asc 1: " & CStr(Asc("1")) & " " & "Asc 9: " & CStr(Asc("9"))
    For i = 1 To Len(Text)
        ch = Mid(Text, i, 1)
        aa = Asc(UCase(ch))
        Select Case aa
        Case Asc("A") To Asc("Z")
        Case Asc("0") To Asc("9")
        Case Else
          Mid(Text, i, 1) = "_"
        End Select
    Next
    ReplNonAlpNum = Text
End Function

'Public Property Get Dock() As ClsDock
'  Set Dock = mDock
'End Property
'Public Property Let Dock(LngVal As ClsDock)
'  mDock.value = LngVal.value
'End Property

Public Property Get Text() As String
    Text = mText
End Property
Public Property Let Text(StrVal As String)
    Dim tci As TCITEM
    mText = StrVal
    If mTCHandle <> 0 Then
        With tci
            .mask = &H1 Or &H2
            '.dwState
            '.dwStateMask
            .pszText = mText
            .cchTextMax = Len(mText)
            .iImage = -1 '.iImage = hIcon
            '.lParam
        End With
        'If index = -1 Then
        'Index = mIndex
        Call SendMessage(mTCHandle, TCM_SETITEM, mIndex, tci)
        If Not mDebugLabel Is Nothing Then mDebugLabel.Caption = mName + mText
    End If
End Property

Public Property Get ImageIndex() As Long
    ImageIndex = mImageIndex
End Property
Public Property Let ImageIndex(LngVal As Long)
    mImageIndex = LngVal
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mToolTipText
End Property
Public Property Let ToolTipText(StrVal As String)
    mToolTipText = StrVal
End Property

Public Property Get Page() As Object 'PictureBox
    Set Page = mPicBPage
End Property
Public Property Set Page(APicB As Object) 'PictureBox)
    Set mPicBPage = APicB
    Call SetParent(mPicBPage.hwnd, mContainer.hwnd)
    mPicBPage.BorderStyle = 0
    mPicBPage.Appearance = 1
    ResizeMe
End Property

Public Sub BringToFront()
    mVisible = True
    mPicBPage.Visible = mVisible
    mPicBPage.ZOrder 0
    ResizeMe
End Sub
Public Sub SendToBack()
    mVisible = False
    'mPicBPage.Visible = mVisible
    mPicBPage.ZOrder 1
End Sub

Public Function ToString() As String
    ToString = "TabPage:" & " " & "{" & mText & "}"
End Function
 
Public Property Get Visible() As Boolean
    'Set Visible = New cBoolean
    Visible = mVisible
End Property

'Public ReadOnly Property Controls
Public Property Get Controls() As ControlCollection
    Set Controls = mControls
End Property

Public Sub ResizeMe()
    Dim L As Single, T As Single, w As Single, H As Single
    L = 0:     T = 0
    w = mContainer.ScaleWidth: H = mContainer.ScaleHeight
    If w > 0 And H > 0 Then Call mPicBPage.Move(L, T, w, H)
End Sub

' #################### '    Private Procedures    ' #################### '
Private Sub mContainer_Resize()
    ResizeMe
End Sub
Private Sub mControls_OnAdding(Ctrl As Object)
    Call SetParent(Ctrl.hwnd, mPicBPage.hwnd)
    Ctrl.Move 0, 0
End Sub

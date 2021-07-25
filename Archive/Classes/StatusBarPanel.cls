VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StatusBarPanel
'          Inherits System.ComponentModel.Component
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt einen Bereich in einem System.Windows.Forms.StatusBar-Steuerelement dar.
Option Explicit
Public Enum HorizontalAlignment 'As Integer
  HorizontalAlignment_Left = 0
  HorizontalAlignment_Right = 1
  HorizontalAlignment_Center = 2
End Enum
Public Enum StatusBarPanelAutoSize 'As Integer
  None = 1
  Spring = 2
  Contents = 3
End Enum
Public Enum StatusBarPanelBorderStyle 'As Integer
  None = 1
  Raised = 2
  Sunken = 3
End Enum
Public Enum StatusBarPanelStyle 'As Integer
  StatusBarPanelStyle_Text = 1
  StatusBarPanelStyle_OwnerDraw = 2
End Enum
Private mAlignment As HorizontalAlignment
Private mAutoSize As StatusBarPanelAutoSize
Private mParent As StatusBar
Private mBorderStyle As Long
Private mMinWidth As Long
Private mStyle As StatusBarPanelStyle
Private mText As String
Private mToolTipText As String
Private mWidth As Long

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  '
End Sub
'Public Sub New()


'Public Property Alignment() As System.Windows.Forms.HorizontalAlignment
Public Property Get Alignment() As HorizontalAlignment
  Alignment = mAlignment
End Property
Public Property Let Alignment(AlignmVal As HorizontalAlignment)
  mAlignment = AlignmVal
End Property

'Public Property AutoSize() As System.Windows.Forms.StatusBarPanelAutoSize
Public Property Get AutoSize() As StatusBarPanelAutoSize

End Property

'Public Overridable NotOverridable Sub BeginInit()
Public Sub BeginInit()

End Sub

'Public Property BorderStyle() As System.Windows.Forms.StatusBarPanelBorderStyle
Public Property Get BorderStyle() As StatusBarPanelBorderStyle
  BorderStyle = mBorderStyle
End Property
Public Property Let BorderStyle(LngVal As StatusBarPanelBorderStyle)
  mBorderStyle = LngVal
End Property

'Public Overridable NotOverridable Sub EndInit()
Public Sub EndInit()

End Sub

'Public Property Icon() As System.Drawing.Icon
Public Property Get Icon() 'As System.Drawing.Icon

End Property
Public Property Let Icon(IcoVal) 'As System.Drawing.Icon

End Property

'Public Property MinWidth() As Integer
Public Property Get MinWidth() As Long
  MinWidth = mMinWidth
End Property
Public Property Let MinWidth(LngVal As Long)
  mMinWidth = LngVal
End Property

'Public ReadOnly Property Parent() As System.Windows.Forms.StatusBar
Public Property Get Parent() As StatusBar
  Set Parent = mParent
End Property

'Public Property Style() As System.Windows.Forms.StatusBarPanelStyle
Public Property Get Style() As StatusBarPanelStyle
  Style = mStyle
End Property
Public Property Let Style(StyleVal As StatusBarPanelStyle)
  mStyle = StyleVal
End Property

'Public Property Text() As String
Public Property Get Text() As String
  Text = mText
End Property
Public Property Let Text(StrVal As String)
  mText = StrVal
End Property

'Public Property ToolTipText() As String
Public Property Get ToolTipText() As String
  ToolTipText = mToolTipText
End Property
Public Property Let ToolTipText(StrVal As String)
  mToolTipText = StrVal
End Property

'Public Overrides Function ToString() As String
Public Function ToString() As String
  ToString = "System.Windows.Forms.StaturBarPanel"
End Function

'Public Property Width() As Integer
Public Property Get Width() As Long
  Width = mWidth
End Property
Public Property Let Width(LngVal As Long)
  mWidth = LngVal
End Property

'Protected Overrides Sub Dispose(ByVal disposing As Boolean)
Public Sub Dispose(ByVal disposing As Boolean)
  
End Sub


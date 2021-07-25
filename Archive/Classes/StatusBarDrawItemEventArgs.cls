VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarDrawItemEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StatusBarDrawItemEventArgs
'          Inherits System.Windows.Forms.DrawItemEventArgs
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt Daten für das System.Windows.Forms.StatusBar.DrawItem-Ereignis bereit.
Option Explicit
Private mPanel As StatusBarPanel

'Public Sub New(ByVal g As System.Drawing.Graphics, ByVal font As System.Drawing.Font, ByVal r As System.Drawing.Rectangle, ByVal itemId As Integer, ByVal itemState As System.Windows.Forms.DrawItemState, ByVal panel As System.Windows.Forms.StatusBarPanel)
'Public Sub New(ByVal g As System.Drawing.Graphics, ByVal font As System.Drawing.Font, ByVal r As System.Drawing.Rectangle, ByVal itemId As Integer, ByVal itemState As System.Windows.Forms.DrawItemState, ByVal panel As System.Windows.Forms.StatusBarPanel, ByVal foreColor As System.Drawing.Color, ByVal backColor As System.Drawing.Color)
Public Sub NewC() 'g As Graphics, font As font, r As Rectangle, itemId As Long, itemState As DrawItemState, panel As StatusBarPanel, Optional foreColor As Color, Optional backColor As Color)

End Sub

'Public ReadOnly Property Panel() As System.Windows.Forms.StatusBarPanel
Public Property Get Panel() As StatusBarPanel
  Set Panel = mPanel
End Property

'Public ReadOnly Property BackColor() As System.Drawing.Color
Public Property Get BackColor() 'As Color

End Property

'Public ReadOnly Property Bounds() As System.Drawing.Rectangle
'Public ReadOnly Property Bounds() As System.Drawing.Rectangle

'Public Overridable Sub DrawBackground()
Public Sub DrawBackground()

End Sub

'Public Overridable Sub DrawFocusRectangle()
Public Sub DrawFocusRectangle()

End Sub

'Public ReadOnly Property Font() As System.Drawing.Font
Public Property Get Font() 'As Font

End Property

'Public ReadOnly Property ForeColor() As System.Drawing.Color
Public Property Get ForeColor() 'As System.Drawing.Color

End Property

'Public ReadOnly Property Graphics() As System.Drawing.Graphics
Public Property Get Graphics() 'As Graphics

End Property

'Public ReadOnly Property Index() As Integer
Public Property Get Index() As Long

End Property

'Public ReadOnly Property State() As System.Windows.Forms.DrawItemState
Public Property Get State() 'As System.Windows.Forms.DrawItemState

End Property


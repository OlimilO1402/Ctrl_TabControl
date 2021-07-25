VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class KeyEventArgs
'          Inherits System.EventArgs
'     Member von: System.Windows.Forms
Option Explicit
Private mKeyData As Long
'Private WithEvents mHandled As cBoolean
Private mHandled As Boolean

Private Sub Class_Initialize()
  '
End Sub

'Public Sub New(ByVal keyData As System.Windows.Forms.Keys)
Public Sub NewC(ByVal KeyData As Long) 'E_Keys)
  mKeyData = KeyData
End Sub
Public Function New_KeyEventArgs(ByVal KeyData As Long) As KeyEventArgs
  Set New_KeyEventArgs = New KeyEventArgs
  Call New_KeyEventArgs.NewC(KeyData)
End Function

'Public Overridable ReadOnly Property Alt() As Boolean
Public Property Get Alt() As Boolean  'ReadOnly
' Ruft einen Wert ab, der angibt, ob ALT gedrückt wurde.
  'Set Alt = New cBoolean
  If mKeyData Then
    Alt = True
  End If
End Property

'Public ReadOnly Property Control() As Boolean
Public Property Get Control() As Boolean  'ReadOnly
  'Set Control = New cBoolean
End Property

'Public Property Handled() As Boolean
Public Property Get Handled() As Boolean
  'Set Handled = New cBoolean
  'set
  Handled = mHandled
End Property
Public Property Let Handled(ByVal bolval As Boolean)
  'Set
  mHandled = bolval
End Property
'Private Sub mHandled_OnLet(ByVal sender As Object)
'  '
'End Sub

'Public ReadOnly Property KeyCode() As System.Windows.Forms.Keys
Public Property Get KeyCode() As Long 'E_Keys  'ReadOnly
  '
End Property

'Public ReadOnly Property KeyData() As System.Windows.Forms.Keys
Public Property Get KeyData() As Long 'E_Keys  'ReadOnly

End Property

'Public ReadOnly Property KeyValue() As Integer
Public Property Get KeyValue() As Long 'cInteger 'ReadOnly

End Property

'Public ReadOnly Property Modifiers() As System.Windows.Forms.Keys
Public Property Get Modifiers() As Long 'E_Keys 'ReadOnly

End Property

'Public Overridable ReadOnly Property Shift() As Boolean
Public Property Get Shift() As Boolean   'ReadOnly
  'Set Shift = New Boolean
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mButtons As Long 'E_MouseButtons 'Left or Right or Middle or XButton1 or XButton2
Private mClicks As Long 'cInteger
Private mDelta As Long 'cInteger
Private mX As Long 'cInteger
Private mY As Long 'cInteger

Private Sub Class_Initialize()
  'Set mClicks = New cInteger
  'Set mX = New cInteger
  'Set mY = New cInteger
  'Set mDelta = New cInteger
End Sub
Public Sub NewC(ByVal Button As Long, ByVal Clicks As Long, ByVal X As Long, ByVal Y As Long, ByVal Delta As Long)
  mButtons = Button
  mClicks = Clicks
  mDelta = Delta
  mX = X
  mY = Y
End Sub
Public Function New_MouseEventArgs(ByVal Button As Long, ByVal Clicks As Long, ByVal X As Long, ByVal Y As Long, ByVal Delta As Long) As MouseEventArgs
  Set New_MouseEventArgs = New MouseEventArgs
  Call New_MouseEventArgs.NewC(Button, Clicks, X, Y, Delta)
End Function
Public Property Get Button() As Long 'E_MouseButtons
  Button = mButtons
End Property
Public Property Get Clicks() As Long 'cInteger
  Clicks = mClicks
End Property
Public Property Get Delta() As Long 'cInteger
  Delta = mDelta
End Property
Public Property Get X() As Long 'cInteger
  X = mX
End Property
Public Property Get Y() As Long 'cInteger
  Y = mY
End Property



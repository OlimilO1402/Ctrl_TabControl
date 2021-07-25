VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarPanelClickEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StatusBarPanelClickEventArgs
'          Inherits System.Windows.Forms.MouseEventArgs
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt Daten für das System.Windows.Forms.StatusBar.PanelClick-Ereignis bereit.
Option Explicit
Private mStatusBarPanel As StatusBarPanel
Private mButton As MouseButtons
Private mClicks As Long
Private mDelta As Long
Private mX As Long
Private mY As Long

Private Sub Class_Initialize()
  '
End Sub
Private Sub Class_Terminate()
  '
End Sub
'Public Sub New(ByVal statusBarPanel As System.Windows.Forms.StatusBarPanel, ByVal button As System.Windows.Forms.MouseButtons, ByVal clicks As Integer, ByVal x As Integer, ByVal y As Integer)
Public Sub NewC(aStatusBarPanel As StatusBarPanel, Button As MouseButtons, Clicks As Long, ByVal X As Long, ByVal Y As Long)
  Set mStatusBarPanel = aStatusBarPanel
  mButton = Button:  mClicks = Clicks:  mX = mX:  mY = Y
End Sub
'Public Sub NewM(abutton As MouseButtons, aclicks As Long, ax As Long, ay As Long, adelta As Long)
'  mButton = abutton: mClicks = aclicks: mX = ax: mY = ay: mDelta = adelta
'End Sub

'Public ReadOnly Property StatusBarPanel() As System.Windows.Forms.StatusBarPanel
Public Property Get StatusBarPanel() As StatusBarPanel
  Set StatusBarPanel = mStatusBarPanel
End Property

'Public ReadOnly Property Button() As System.Windows.Forms.MouseButtons
Public Property Get Button() As MouseButtons
  Button = mButton
End Property

'Public ReadOnly Property Clicks() As Integer
Public Property Get Clicks() As Long
  Clicks = mClicks
End Property

'Public ReadOnly Property Delta() As Integer
Public Property Get Delta() As Long
  Delta = mDelta
End Property

'Public ReadOnly Property X() As Integer
Public Property Get X() As Long
  X = mX
End Property

'Public ReadOnly Property Y() As Integer
Public Property Get Y() As Long
  Y = mY
End Property

'Dim sbpcea As New StatusBarPanelClickEventArgs
'sbpcea.Button()
'sbpcea.Clicks()
'sbpcea.Delta()
'sbpcea.Empty()
'sbpcea.GetType()
'sbpcea.StatusBarPanel()
'sbpcea.X()
'sbpcea.Y()



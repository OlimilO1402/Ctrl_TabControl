VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StatusBar
'          Inherits System.Windows.Forms.Control
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt ein Statusleiste-Steuerelement von Windows dar.
Option Explicit
Private Const WC_STATUSBAR As String = "msctls_statusbar32"
Private Const WC_STATUSBARA As String = "msctls_statusbar32"
Private Const WC_STATUSBARW As String = "msctls_statusbar32"
'"msctls_statusbar32"
'"msctls_progress32"

Private mOwner As PictureBox
Private mShowPanels As Boolean
Private mPanels As StatusBarPanelCollection
Private mSizingGrip As Boolean
Private mTabStop As Boolean
Private mText As String
Public Event DrawItem(ByVal sender As Object, ByVal sbdevent As StatusBarDrawItemEventArgs)
Public Event PanelClick(ByVal sender As Object, ByVal e As StatusBarPanelClickEventArgs)

Private Sub Class_Initialize()
  Set mPanels = New StatusBarPanelCollection
End Sub
Private Sub Class_Terminate()
  Set mPanels = Nothing
End Sub

'Public Sub New()
Private Sub Create(myOwner As PictureBox)
  Set mOwner = myOwner
End Sub

'Public Overrides Property BackColor() As System.Drawing.Color
Public Property Get BackColor() As Long 'System.Drawing.Color

End Property
Public Property Let BackColor(acolor As Long) 'System.Drawing.Color

End Property

'Public Overrides Property Dock() As System.Windows.Forms.DockStyle
Public Property Get Dock() As E_DockStyle 'System.Windows.Forms.DockStyle

End Property

'Public Overrides Property Font() As System.Drawing.Font
Public Property Get Font() As StdFont 'System.Drawing.Font

End Property
Public Property Let Font(FntVal As StdFont) 'System.Drawing.Font

End Property

'Public ReadOnly Property Panels() As System.Windows.Forms.StatusBar.StatusBarPanelCollection
Public Property Get panels() As StatusBarPanelCollection
  
End Property

'Public Property ShowPanels() As Boolean
Public Property Get ShowPanels() As Boolean
  ShowPanels = mShowPanels
End Property
Public Property Let ShowPanels(bolval As Boolean)
  mShowPanels = bolval
End Property

'Public Property SizingGrip() As Boolean
Public Property Get SizingGrip() As Boolean
  SizingGrip = mSizingGrip
End Property
Public Property Let SizingGrip(bolval As Boolean)
  mSizingGrip = bolval
End Property

'Public Property TabStop() As Boolean
Public Property Get TabStop() As Boolean
  TabStop = mTabStop
End Property
Public Property Let TabStop(bolval As Boolean)
  mTabStop = bolval
End Property

'Public Overrides Property Text() As String
Public Property Get Text() As String
  Text = mText
End Property
Public Property Let Text(StrVal As String)
  mText = StrVal
End Property

'Public Overrides Function ToString() As String
Public Function ToString() As String
  '
End Function

'Protected Overrides Sub CreateHandle()
Public Sub CreateHandle()

End Sub

'Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
Public Property Get CreateParams() 'As CreateParams

End Property

'Protected Overrides ReadOnly Property DefaultImeMode() As System.Windows.Forms.ImeMode
Public Property Get DefaultImeMode() 'As ImeMode

End Property

'Protected Overrides ReadOnly Property DefaultSize() As System.Drawing.Size
Public Property Get DefaultSize() 'As Size

End Property

'Protected Overrides Sub Dispose(ByVal disposing As Boolean)
Public Sub Dispose(ByVal disposing As Boolean)

End Sub

'Protected Overridable Sub OnDrawItem(ByVal sbdievent As System.Windows.Forms.StatusBarDrawItemEventArgs)
Public Sub OnDrawItem(ByVal sbdievent As StatusBarDrawItemEventArgs)
  Dim ee As New StatusBarDrawItemEventArgs
  'call ee.NewC
  RaiseEvent DrawItem(sbdievent, ee)
End Sub



'
'StatusBar1.AccessibleDescription()
'StatusBar1.AccessibleName()
'StatusBar1.AccessibleRole()
'StatusBar1.AllowDrop()
'StatusBar1.Anchor()
'StatusBar1.BackColor()
'StatusBar1.BackgroundImage()
'StatusBar1.BringToFront()
'StatusBar1.CausesValidation()
'StatusBar1.Container()
'StatusBar1.Contains()
'StatusBar1.ContextMenu()
'StatusBar1.ControlAccessibleObject()
'StatusBar1.ControlCollection()
'StatusBar1.Controls()
'StatusBar1.CreateControl()
'StatusBar1.CreateGraphics()
'StatusBar1.CreateObjRef()
'StatusBar1.Cursor()
'StatusBar1.DataBindings()
'StatusBar1.DefaultBackColor()
'StatusBar1.DefaultFont()
'StatusBar1.DefaultForeColor()
'StatusBar1.Dispose()
'StatusBar1.Dock()
'StatusBar1.DoDragDrop()
'StatusBar1.Enabled()
'StatusBar1.FindForm()
'StatusBar1.Focus()
'StatusBar1.Font()
'StatusBar1.ForeColor()
'StatusBar1.GetChildAtPoint()
'StatusBar1.GetContainerControl()
'StatusBar1.GetLifetimeService()
'StatusBar1.GetNextControl()
'StatusBar1.GetType()
'StatusBar1.Handle()
'StatusBar1.Height()
'StatusBar1.Hide()
'StatusBar1.ImeMode()
'StatusBar1.InitializeLifetimeService()
'StatusBar1.Invalidate()
'StatusBar1.Invoke()
'StatusBar1.IsMnemonic()
'StatusBar1.Left()
'StatusBar1.Location()
'StatusBar1.ModifierKeys()
'StatusBar1.MouseButtons()
'StatusBar1.MousePosition()
'StatusBar1.Name()
'StatusBar1.Panels()
'StatusBar1.Parent()
'StatusBar1.PointToClient()
'StatusBar1.PointToScreen()
'StatusBar1.PreProcessMessage()
'StatusBar1.RectangleToClient()
'StatusBar1.RectangleToScreen()
'StatusBar1.Refresh()
'StatusBar1.ResetText()
'StatusBar1.ResumeLayout()
'StatusBar1.RightToLeft()
'StatusBar1.Scale()
'StatusBar1.Select()
'StatusBar1.SelectNextControl()
'StatusBar1.SendToBack()
'StatusBar1.SetBounds()
'StatusBar1.Show()
'StatusBar1.ShowPanels()
'StatusBar1.Size()
'StatusBar1.SizingGrip()
'StatusBar1.StatusBarPanelCollection()
'StatusBar1.SuspendLayout()
'StatusBar1.TabIndex()
'StatusBar1.TabStop()
'StatusBar1.Tag()
'StatusBar1.Text()
'StatusBar1.Top()
'StatusBar1.ToString()
'StatusBar1.Update()
'StatusBar1.Visible()
'StatusBar1.Width()


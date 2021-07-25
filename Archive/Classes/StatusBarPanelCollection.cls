VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StatusBarPanelCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class StatusBar.StatusBarPanelCollection
'          Inherits System.Object
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt die Auflistung der Bereiche in einem System.Windows.Forms.StatusBar-Steuerelement dar.
Option Explicit
Private mCol As Collection
Private mOwner As StatusBar
Private mIsReadOnly As Boolean

Private Sub Class_Initialize()
  Clear
End Sub
Private Sub Class_Terminate()
  Set mCol = Nothing
End Sub
'Public Sub New(ByVal owner As System.Windows.Forms.StatusBar)
Public Sub NewC(ByVal owner As StatusBar)
  Set mOwner = owner
End Sub

'Public Overridable Function Add(ByVal text As String) As System.Windows.Forms.StatusBarPanel
Public Function Add(Text As String) As StatusBarPanel
  Set Add = New StatusBarPanel
  Add.Text = Text
End Function
Public Function AddI(ByVal value As StatusBarPanel) As Long
  Call mCol.Add(value)
End Function
'Public Overridable Sub AddRange(ByVal panels() As System.Windows.Forms.StatusBarPanel)
Public Sub AddRange(panels() As StatusBarPanel)
Dim stbp 'As StatusBarPanel
  For Each stbp In panels
    Call mCol.Add(stbp)
  Next
End Sub

'Public Overridable Sub Clear()
Public Sub Clear()
  Set mCol = New Collection
End Sub

'Public Function Contains(ByVal panel As System.Windows.Forms.StatusBarPanel) As Boolean
Public Function Contains(ByVal panel As StatusBarPanel) As Boolean
Dim stbp As StatusBar
  For Each stbp In mCol
    If stbp Is panel Then
      Contains = True
      Exit Function
    End If
  Next
End Function

'Public Function GetEnumerator() As System.Collections.IEnumerator
Public Function GetEnumerator() As IUnknown 'System.Collections.IEnumerator
Attribute GetEnumerator.VB_UserMemId = -4
  Set GetEnumerator = mCol.[_NewEnum]
End Function

'Public Function IndexOf(ByVal panel As System.Windows.Forms.StatusBarPanel) As Integer
Public Function IndexOf(ByVal panel As StatusBarPanel) As Long
Dim i As Long, stbp As StatusBarPanel
  For i = 1 To mCol.Count
    If mCol(i) Is panel Then
      IndexOf = i - 1
      Exit For
    End If
  Next
End Function

'Public Overridable Sub Insert(ByVal index As Integer, ByVal value As System.Windows.Forms.StatusBarPanel)
Public Sub Insert(ByVal index As Long, ByVal value As StatusBarPanel)
  Call mCol.Add(value, index)
End Sub

'Public Overridable NotOverridable ReadOnly Property IsReadOnly() As Boolean
Public Property Get IsReadonly() As Boolean
  IsReadonly = mIsReadOnly
End Property
Friend Property Let IsReadonly(bolval As Boolean)
  mIsReadOnly = bolval
End Property

'Public Overridable Default Property Item(ByVal index As Integer) As System.Windows.Forms.StatusBarPanel
Public Property Get Item(ByVal index As Long) As StatusBarPanel
Attribute Item.VB_UserMemId = 0
  If index + 1 <= mCol.Count Then
    Set Item = mCol(index + 1)
  End If
End Property

'Public Overridable Sub Remove(ByVal value As System.Windows.Forms.StatusBarPanel)
Public Sub Remove(ByVal value As StatusBarPanel)
Dim i As Long
  For i = 1 To mCol.Count
    If mCol(i) Is value Then
      Call mCol.Remove(i)
    End If
  Next
End Sub

'Public Overridable Sub RemoveAt(ByVal index As Integer)
Public Sub RemoveAt(ByVal index As Long)
'mCol ist 1-basiert
  'deshalb bei index eins dazunehmen
  If index + 1 <= mCol.Count Then
    mCol.Remove (index)
  End If
End Sub


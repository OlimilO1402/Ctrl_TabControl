VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ClsInteger"
Attribute VB_Ext_KEY = "Member0" ,"ClsInteger"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Public Class Control.ControlCollection
'          Inherits System.Object
'     Member von: System.Windows.Forms
'Zusammenfassung:
' Stellt eine Auflistung von System.Windows.Forms.Control-Objekten dar.
Option Explicit
'Klasse ist Bestandteil des .NET-Frameworks
Private mControlCol As Collection
Public IsReadOnly As Boolean
Public Event OnAdding(Ctrl As Object)

Private Sub Class_Initialize()
  IsReadOnly = False
  Clear
End Sub

Public Sub Add(Ctrl As Object)
  If Not IsReadOnly Then mControlCol.Add Ctrl
  RaiseEvent OnAdding(Ctrl)
  'jetzt auch den Parent setzen
End Sub
Public Sub AddRange(Ctrl() As Object)
Dim i As Long
  If Not IsReadOnly Then
    For i = LBound(Ctrl()) To UBound(Ctrl())
      Add (Ctrl(i))
    Next
  End If
End Sub

Public Sub Clear()
  If Not IsReadOnly Then
    Set mControlCol = New Collection
  End If
End Sub
Public Function Contains(Ctrl As Object) As Boolean
Dim C As Object
  For Each C In mControlCol
    If C Is Ctrl Then
      Contains = True
      Exit Function
    End If
  Next
End Function
Public Sub CopyTo(Ctrl() As Object, Index As Long)
'.NET: Kopiert die gesamte Auflistung an eine angegebene Position im vorhandenen Array
  'was wird hier gemacht?
  'handelt es sich bei diesem Array um eine Liste oder um ein dynamisches Array?
  'werden alle Elemente in das Array eingefügt,
  'so daß vorhandene Positionen nach index nach hinten geschoben werden, oder
  'so daß vorhandene Positionen nach index überschrieben werden
  'was ist wenn das Array nicht die erforderliche Größe hat?
End Sub
Public Property Get Count() As Long
  Count = mControlCol.Count
End Property

Public Function GetChildIndex(Ctrl As Object) As Long
'.NET: ruft den Index des angegebenen untergeordneten Steuerelements in der
'Steuerelementauflistung ab.
  'was wird hier gemacht? ist das nicht das gleich wie IndexOf?
End Function

Public Sub SetChildIndex(child As Object, newindex As Long)
'.NET: legt den Index des angegebenen untergeordneten Steuerelements in der
'Auflistung auf den angegebenen Indexwert fest.
  'was wird hier gemacht?
End Sub

Public Property Get Item(Index As Long) As Object
Attribute Item.VB_UserMemId = 0
  If Count > 0 Then
    If Index < 0 Then Index = 0
    Set Item = mControlCol(Index + 1)
  End If
End Property

Public Function IndexOf(Ctrl As Object) As Long
Dim C As Object
  For Each C In mControlCol
    If C Is Ctrl Then Exit Function
    IndexOf = IndexOf + 1
  Next
End Function

'Public Property Get IsReadOnly() As Boolean
'  IsReadOnly = mReadOnly
'End Property
'Public Property Let IsReadOnly(BolVal As Boolean)
'  mReadOnly = BolVal
'End Property

Public Sub RemoveAt(Index As Long)
  If Not IsReadOnly Then Call mControlCol.Remove(Index)
End Sub
Public Sub Remove(Ctrl As Object)
  If Not IsReadOnly Then Call RemoveAt(IndexOf(Ctrl))
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mControlCol.[_NewEnum]
End Property

Private Sub Class_Terminate()
  Set mControlCol = Nothing
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DockStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Public Enum DockStyle As Integer
'     Member von: System.Windows.Forms
Option Explicit
'Private mEnum As New cEnum
Private mValue As E_DockStyle 'im Modul
Attribute mValue.VB_VarMemberFlags = "40"
Attribute mValue.VB_VarUserMemId = 0

Public Enum E_DockStyle
  DockStyleNone = 0
  DockStyleTop = 1
  DockStyleBottom = 2
  DockStyleLeft = 3
  DockStyleRight = 4
  DockStyleFill = 5
End Enum

Public Property Get value() As E_DockStyle
Attribute value.VB_UserMemId = 0
Attribute value.VB_MemberFlags = "40"
  value = mValue
End Property
Public Property Let value(dsVal As E_DockStyle)
  mValue = dsVal
End Property

'Public Shared Const None As System.Windows.Forms.DockStyle = 0
Public Property Get None() As DockStyle
  Set None = New DockStyle
  None = E_DockStyle.DockStyleNone
End Property

'Public Shared Const Top As System.Windows.Forms.DockStyle = 1
Public Property Get Top() As DockStyle
  Set Top = New DockStyle
  Top = E_DockStyle.DockStyleTop
End Property

'Public Shared Const Bottom As System.Windows.Forms.DockStyle = 2
Public Property Get Bottom() As DockStyle
  Set Bottom = New DockStyle
  Bottom = E_DockStyle.DockStyleBottom
End Property

'Public Shared Const Left As System.Windows.Forms.DockStyle = 3
Public Property Get Left() As DockStyle
  Set Left = New DockStyle
  Left = E_DockStyle.DockStyleLeft
End Property

'Public Shared Const Right As System.Windows.Forms.DockStyle = 4
Public Property Get Right() As DockStyle
  Set Right = New DockStyle
  Right = E_DockStyle.DockStyleRight
End Property

'Public Shared Const Fill As System.Windows.Forms.DockStyle = 5
Public Property Get Fill() As DockStyle
  Set Fill = New DockStyle
  Fill = E_DockStyle.DockStyleFill
End Property

Public Function CompareTo(ObjA As Object, Optional ObjB As Object) As Long 'cInteger
'  Set CompareTo = New cInteger
'  If ObjA = ObjB Then
'    CompareTo = 0
'  Else
'    If ObjA < ObjB Then
'      CompareTo = -1
'    Else
'      CompareTo
'    End If
'  End If
  'Set CompareTo = mEnum.CompareTo(ObjA)
End Function

Public Function Equals(obj As Object) As Boolean 'cBoolean
  'Set Equals = New cBoolean
  If obj = Me Then Equals = True
End Function

Public Function Format(value As Object, frmt As String)
  '
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnum"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Public MustInherit Class Enum
Option Explicit
'Implements IComparable
'Implements IConvertible
'Implements IFormattable
'diese Klasse ist hier eigentlich nur Prototyp
'jedes Enum muﬂ eine eigene Klasse bilden
'Dies ist die Enum-Klasse zu .NET
'Public Value As Long
'
'Private Function IComparable_CompareTo(ByVal obj As Object) As Long 'cInteger
'  Set IComparable_CompareTo = CompareTo(obj)
'End Function
'Public Overridable NotOverridable Function CompareTo(ByVal target As Object) As Integer
'Public Function CompareTo(ByVal target As Object) As Long 'cInteger
'  If TypeOf target Is cEnum Then
'
'  End If
'End Function

'Public Overrides Function Equals(ByVal obj As Object) As Boolean
'Public Function Equals(ByVal obj As Object) As Boolean 'cBoolean
'  '
'End Function

'Public Shared Function Format(ByVal enumType As System.Type, ByVal value As Object, ByVal format As String) As String
'Public Static Function fFormat(ByVal enumType As SystemType, ByVal ObjValue As Object, ByVal sFormat As String) As String 'cString
'  '
'End Function

'Public Shared Function GetName(ByVal enumType As System.Type, ByVal value As Object) As String
'Public Static Function GetName(ByVal enumType As SystemType, ByVal ObjValue As Object) As String 'cString
'  '
'End Function

'Public Shared Function GetNames(ByVal enumType As System.Type) As String()
'Public Static Function GetNames(ByVal enumType As SystemType) As Variant 'String()
'  '
'End Function

'Private Function IConvertible_GetTypeCode() As SystemTypeCode
'  Set IConvertible_GetTypeCode = GetTypeCode
'End Function
'Public Overridable NotOverridable Function GetTypeCode() As System.TypeCode
'Public Function GetTypeCode() 'As SystemTypeCode
'  '
'End Function

'Public Shared Function GetUnderlyingType(ByVal enumType As System.Type) As System.Type
'Public Static Function GetUnderlyingType(ByVal enumType As SystemType) 'As SystemType
'  '
'End Function

'Public Shared Function GetValues(ByVal enumType As System.Type) As System.Array
'Public Static Function GetValues(ByVal enumType As SystemType) 'As SystemArray
'  '
'End Function

'Public Shared Function IsDefined(ByVal enumType As System.Type, ByVal value As Object) As Boolean
'Public Static Function IsDefined(ByVal enumType As SystemType, ByVal ObjValue As Object) As Boolean 'cBoolean
'  '
'End Function

'Public Shared Function Parse(ByVal enumType As System.Type, ByVal value As String) As Object
'Public Shared Function Parse(ByVal enumType As System.Type, ByVal value As String, ByVal ignoreCase As Boolean) As Object
'Public Function Parse(ByVal enumType As SystemType, ByVal StrVal As String, Optional ByVal ignoreCase As Boolean) As Object
'  '
'End Function

'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As Byte) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As Integer) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As Long) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As Object) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As Object) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As System.SByte) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As System.UInt16) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As System.UInt32) As Object
'Public Shared Function ToObject(ByVal enumType As System.Type, ByVal value As System.UInt64) As Object
'Public Function ToObject(ByVal enumType As SystemType, ByVal Value As Variant) As Object
'  '
'End Function

'Private Function IConvertible_ToString(ByVal provider As System.IFormatProvider) As String 'cString
'  Set IConvertible_ToString = ToString(, provider)
'End Function
'Private Function IFormattable_ToString(ByVal sFormat As String, ByVal formatProvider As IFormatProvider) As cString
'  Set IFormattable_ToString = ToString(sFormat, formatProvider)
'End Function
'Public Overrides Function ToString() As String
'Public Function ToString(ByVal format As String) As String
'Public Overridable NotOverridable Function ToString(ByVal provider As System.IFormatProvider) As String
'Public Overridable NotOverridable Function ToString(ByVal provider As System.IFormatProvider) As String
'Public Function ToString(Optional ByVal sFormat As String, Optional ByVal provider As System.IFormatProvider) As String 'cString
'  Set ToString = New cString
'  'Select Case mValue
'  'End Select
'End Function

'################'  v     Implements IConvertible      v   '################'
'Private Function IConvertible_ToBoolean(ByVal provider As IFormatProvider) As Boolean 'cBoolean
'
'End Function
'
'Private Function IConvertible_ToByte(ByVal provider As IFormatProvider) As Byte 'clByte
'
'End Function
'
'Private Function IConvertible_ToChar(ByVal provider As IFormatProvider) As String 'cChar
'
'End Function
'
'Private Function IConvertible_ToDateTime(ByVal provider As IFormatProvider) As Date
'
'End Function
'
'Private Function IConvertible_ToDouble(ByVal provider As IFormatProvider) As Double 'cDouble
'
'End Function
'
'Private Function IConvertible_ToInt16(ByVal provider As IFormatProvider) As Integer 'cInt16
'
'End Function
'
'Private Function IConvertible_ToInt32(ByVal provider As IFormatProvider) As Long 'cInteger
'
'End Function
'
'Private Function IConvertible_ToSByte(ByVal provider As IFormatProvider) As Byte 'SByte
'
'End Function
'
'Private Function IConvertible_ToSingle(ByVal provider As IFormatProvider) As Single 'cSingle
'
'End Function
'

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Class_Initialize()
  '
End Sub

Public Sub DDoEvents()
  DoEvents
End Sub

Public Sub eExit()
Dim Frm As Form
  For Each Frm In VB.Forms
    Unload Frm
  Next
End Sub

Public Sub EnableVisualStyles()
  InitCommonControls
End Sub



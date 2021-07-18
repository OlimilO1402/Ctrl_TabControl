VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------------------------
'cSubclass - module-less, IDE safe, machine code window subclasser
'
'v1.00 20030107 First cut..........................................................................
'
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'no longer need for the subclass.tlb
Private Const OFFSET_P1   As Long = 9       'Callback gate address
Private Const OFFSET_P2   As Long = 22      'Before table entry count
Private Const OFFSET_P3   As Long = 37      'Before table address
Private Const OFFSET_P4   As Long = 55      'Close gate value
Private Const OFFSET_P5   As Long = 84      'Owner object address for iSubclass_Before
Private Const OFFSET_P6   As Long = 121     'Original WndProc address
Private Const OFFSET_P7   As Long = 126     'CallWindowProc address
Private Const OFFSET_P8   As Long = 134     'After table entry count
Private Const OFFSET_P9   As Long = 149     'After table address
Private Const OFFSET_PA   As Long = 167     'Close gate value
Private Const OFFSET_PB   As Long = 188     'Owner object address for iSubclass_After
Private Const ARRAY_LB    As Long = 1       'Lowest bound of arrays. Don't change to
                                            'zero else the AddMsg logic will fail
Private Type tCode
    Buf(ARRAY_LB To 213)  As Byte           'Code buffer
End Type

Private Type tCodeBuf
    code                  As tCode          'Subclass WndProc code
End Type

Private CodeBuf           As tCodeBuf       'Subclass WndProc code instance
Private nBreakGate        As Long           'Callback breakpoint gate
Private nMsgCntB          As Long           'Before msg table entry count
Private nMsgCntA          As Long           'After msg table entry count
'Achtung am Besten die Subclass.Tlb ins Verzeichnis kopieren
'SubClass.tlb wird jetzt nicht mehr gebraucht -> durch iSubClass
Private aMsgTblB()        As WM_ENUM        'Before msg table array 'in Modul
Private aMsgTblA()        As WM_ENUM        'After  msg table array
Private hWndSubclass      As Long           'Handle of the window being subclassed
Private nWndProcSubclass  As Long           'The address of our WndProc
Private nWndProcOriginal  As Long           'The address of the existing WndProc
Private nGateValue        As Long           'Callback breakpoint gate value

'-----------------------------
' Private subroutines

'Worker sub for AddMsg
Private Sub AddMsgSub(uMsg As WM_ENUM, aMsgTbl() As WM_ENUM, nMsgCnt As Long, When As MSG_WHEN_ENUM)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                           'If ALL_MESSAGES
        nMsgCnt = -1                                      'Indicates that all messages are to callback
    Else                                                  'Else a specific message number
        For nEntry = ARRAY_LB To nMsgCnt                  'For each existing entry. NB will skip if 0 or -1 (ALL_MESSAGES)
            Select Case aMsgTbl(nEntry)                   'Select on the message number stored in this table entry
                Case -1                                   'This msg table slot is a deleted entry
                    aMsgTbl(nEntry) = uMsg                'Re-use this entry
                    Exit Sub                              'Bail
                Case uMsg                                 'The msg is already in the table!
                    Exit Sub                              'Bail
            End Select
        Next nEntry                                       'Next entry

        'Make space for the new entry
        ReDim Preserve aMsgTbl(ARRAY_LB - 1 To nEntry)    'Increase the size of the table. NB nEntry = nMsgCnt + 1
        nMsgCnt = nEntry                                  'Bump the entry count
        aMsgTbl(nEntry) = uMsg                            'Store the message in the table
    End If
    
    If hWndSubclass <> 0 Then                             'If the subclasser is active
        If When = MSG_BEFORE Then                         'If before
            Call PatchVal(OFFSET_P2, nMsgCnt)             'Patch the Before table entry count
            Call PatchVal(OFFSET_P3, AddrMsgTbl(aMsgTbl)) 'Patch the address of the Before message table. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.
        Else                                              'Else after
            Call PatchVal(OFFSET_P8, nMsgCnt)             'Patch the After table entry count
            Call PatchVal(OFFSET_P9, AddrMsgTbl(aMsgTbl)) 'Patch the address of the After message table. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.
        End If
    End If
End Sub

'Return the address of the CallWindowProc function
Private Function AddrCallWindowProc() As Long
    AddrCallWindowProc = GetProcAddress(GetModuleHandle("user32"), "CallWindowProcA")
End Function

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl(aMsgTbl() As WM_ENUM) As Long
  '  On Error Resume Next                                 'The table may not be dimensioned yet so we need protection
    If UBound(aMsgTbl()) > 0 Then AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB)) 'Get the address of the first element of the passed message table
  '  On Error GoTo 0                                      'Switch off error protection
End Function

'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
    Dim i     As Long
    Dim j     As Long
    
    ReDim aMsgTblA(0)
    ReDim aMsgTblB(0)
    
    Const OPS As String = _
      "558BEC83C4F85756BE_patch1_33C08945FC8945F8B9_patch2_83F900744D83F9FF740CBF_patch3_8B450CF2AF753C833E007537C" & _
      "706_patch4_8D4514508D4510508D450C508D4508508D45FC508D45F850BA_patch5_528B02FF5020C706000000008B45F883F80075" & _
      "60FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB9_patch8_83F900743D83F9FF740CBF_patch9_8B450CF2AF752C8" & _
      "33E007527C706_patchA_FF7514FF7510FF750CFF75088D45FC50BA_patchB_528B02FF501CC706000000005E5F8B45FCC9C2100090"
    
    'Convert the string of opcodes from hex pairs to bytes and store in the code buffer
    With CodeBuf.code
        j = 1                                             'Set the character index to the start of the opcode string
        For i = ARRAY_LB To UBound(.Buf)                  'For each byte of the code buufer
            .Buf(i) = Val("&H" & Mid$(OPS, j, 2))         'Pull a pair of hex characters and convert to a byte
            j = j + 2                                     'Bump to the next pair of characters
        Next i                                            'Next byte of the code buffer

        nWndProcSubclass = VarPtr(.Buf(ARRAY_LB))         'Address of the Subclass WndProc entry point
    End With
    nGateValue = InIDE                                    'Set the gate value so that the callback gate is shut whilst running in the IDE but not in the compiled app
End Sub

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
    If hWndSubclass <> 0 Then Call UnSubclass             'UnSubclass if the subclasser is active
End Sub

'Worker sub for DelMsg
Private Sub DelMsgSub(uMsg As WM_ENUM, aMsgTbl() As WM_ENUM, nMsgCnt As Long, When As MSG_WHEN_ENUM)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                           'If deleting all messages (specific or ALL_MESSAGES)
        nMsgCnt = 0                                       'Message count is now zero

        If hWndSubclass <> 0 Then                         'If the subclasser is active
            If When = MSG_BEFORE Then                     'If before
                Call PatchVal(OFFSET_P2, 0)               'Patch the before table message count
            Else                                          'Else after
                Call PatchVal(OFFSET_P8, 0)               'Patch the after table message count
            End If
        End If
    Else                                                  'Else deleteting a specific message
        For nEntry = ARRAY_LB To nMsgCnt                  'For each table entry
            If aMsgTbl(nEntry) = uMsg Then                'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = -1                      'Mark the table slot as available
                Exit For                                  'Bail
            End If
        Next nEntry                                       'Next entry
    End If
End Sub

'Return -1 if we're running in the IDE or 0 if were running compiled
Private Function InIDE() As Long
    Static value As Long
    
    If value = 0 Then
        value = 1
        Debug.Assert InIDE() Or True 'This line won't exist in the compiled app
        InIDE = value - 1
    End If
    
    value = 0
End Function

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
    Call CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)
End Sub

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
    Call CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)
End Sub

'-----------------------------
'Public interface

'Call this method to add a message to the msg callback table. NB This method can be called at any time
Public Sub AddMsg(uMsg As WM_ENUM, When As MSG_WHEN_ENUM)
    If When = MSG_BEFORE Then                             'If before
        Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)    'Add the message, pass the before table and before table message count variables
    Else                                                  'Else after
        Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)    'Add the message, pass the after table and after table message count variables
    End If
End Sub

'Call the original WndProc
Public Function CallOrigWndProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hWndSubclass <> 0 Then
        CallOrigWndProc = CallWindowProc(nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam)
    Else
        Debug.Assert False                                'LOGIC ERROR: The subclasser isn't active!
    End If
End Function

'Call this method to delete a message from the msg table. NB This method can be called at any time
Public Sub DelMsg(uMsg As WM_ENUM, When As MSG_WHEN_ENUM)
    If When = MSG_BEFORE Then                             'If before
        Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)    'Delete the message, pass the before table and before message count variables
    Else                                                  'Else after
        Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)    'Delete the message, pass the after table and after message count variables
    End If
End Sub

'Call this method to subclass the passed window handle
Public Sub Subclass(hhwnd As Long, owner As iSubClass)
    Debug.Assert (hWndSubclass = 0)                       'LOGIC ERROR: The subclasser is already active!
    
    hWndSubclass = hhwnd                                  'Store the window handle
    nWndProcOriginal = GetWindowLong(hhwnd, GWL_WNDPROC)  'Address of the original WndProc
    
    Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))          'Breakpoint gate address
    Call PatchVal(OFFSET_P2, nMsgCntB)                    'Before table entry count
    Call PatchVal(OFFSET_P3, AddrMsgTbl(aMsgTblB))        'Before table address
    Call PatchVal(OFFSET_P4, nGateValue)                  'Callback gate (close if in IDE)
    Call PatchVal(OFFSET_P5, ObjPtr(owner))               'Owner object address for iSubclass_Before
    Call PatchVal(OFFSET_P6, nWndProcOriginal)            'Original WndProc address
    Call PatchRel(OFFSET_P7, AddrCallWindowProc)          'Address of CallWindowProc
    Call PatchVal(OFFSET_P8, nMsgCntA)                    'After table entry count
    Call PatchVal(OFFSET_P9, AddrMsgTbl(aMsgTblA))        'After table address
    Call PatchVal(OFFSET_PA, nGateValue)                  'Callback gate (close if in IDE)
    Call PatchVal(OFFSET_PB, ObjPtr(owner))               'Owner object address for iSubclass_After
    Call SetWindowLong(hhwnd, GWL_WNDPROC, nWndProcSubclass) 'Set our WndProc in place of the original
End Sub

'Call this method to stop subclassing the window
Public Sub UnSubclass()
    If hWndSubclass <> 0 Then
        Call PatchVal(OFFSET_P2, 0)                         'Patch the code to ensure no further iSubclass_Before callbacks
        Call PatchVal(OFFSET_P8, 0)                         'Patch the code to ensure no further iSubclass_After callbacks
        Call SetWindowLong(hWndSubclass, _
                           GWL_WNDPROC, nWndProcOriginal)   'Restore the original WndProc
        hWndSubclass = 0                                    'Indicate the subclasser is inactive
        nMsgCntB = 0                                        'Message before count zero
        nMsgCntA = 0                                        'Message after count zero
    End If
End Sub



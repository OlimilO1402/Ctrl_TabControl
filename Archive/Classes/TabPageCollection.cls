VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabPageCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Collection für ClsTabPage
Option Explicit
Private Const TCM_FIRST             As Long = &H1300  '{ Tab control messages }   Dez: 4864

Private Const TCM_GETITEMA          As Long = (TCM_FIRST + 5)
Private Const TCM_SETITEMA          As Long = (TCM_FIRST + 6)
Private Const TCM_INSERTITEMA       As Long = (TCM_FIRST + 7)
Private Const TCM_DELETEITEM        As Long = (TCM_FIRST + 8)

Private Const TCM_GETITEM           As Long = TCM_GETITEMA
Private Const TCM_SETITEM           As Long = TCM_SETITEMA
Private Const TCM_INSERTITEM       As Long = TCM_INSERTITEMA
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private mParent As Form           'das Formular in dem alle TabPages erstellt werden
Private WithEvents mContainer As PictureBox 'mein Container
Attribute mContainer.VB_VarHelpID = -1
Private WithEvents mPageCont As PictureBox  'Container für alle TabPages
Attribute mPageCont.VB_VarHelpID = -1
Private mTCName As String
Private mHandle As Long
Private mPageCol As Collection    'Liste der TabPages

Private Sub Class_Initialize()
  Call Clear
End Sub
Public Sub Create(myOwner As Object, MyContainer As Object, MyPageCont As Object, MyTCName As String, MyTCHandle As Long)
  Set mParent = myOwner 'Form
  Set mContainer = MyContainer 'Picturebox
  Set mPageCont = MyPageCont 'PictureBox
  mTCName = MyTCName
  mHandle = MyTCHandle
  'mPageCont.BorderStyle = 0
  mPageCont.ZOrder 0
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mPageCol.[_NewEnum]
End Property

'Public Property Get Handle() As Long
'  Handle = mHandle
'End Property
'Public Property Let Handle(hWnd As Long)
'  mHandle = hWnd
'End Property
Public Sub Add(NewTabPage As TabPage)
Dim udtTCItem As TCITEM
Dim index As Long
Dim TabHndl As Long
  If Not mPageCont.Visible Then mPageCont.Visible = True

  NewTabPage.mTCHandle = mHandle
  Call NewTabPage.Create(mParent, mPageCont, mTCName, NewTabPage.Text)
  mPageCol.Add NewTabPage
  With udtTCItem
    .mask = &H1 Or &H2
    '.dwState
    '.dwStateMask
    .pszText = NewTabPage.Text
    .cchTextMax = Len(NewTabPage.Text)
    .iImage = -1 '.iImage = hIcon
    '.lParam
  End With
  'If index = -1 Then
  index = Count - 1
  NewTabPage.mIndex = index
  Call SendMessage(mHandle, TCM_INSERTITEMA, index, udtTCItem)
  'MsgBox CStr(TabHndl)
End Sub
Public Sub AddRange(pages() As TabPage)
Dim i As Long
  For i = LBound(pages) To UBound(pages)
    mPageCol.Add (pages(i))
  Next
End Sub

Public Sub Clear()
  Set mPageCol = New Collection
End Sub
Public Property Get Count() As Long 'cInteger
  'Set Count = New cInteger
  Count = mPageCol.Count
End Property
Public Function Contains(Page As TabPage) As Boolean
Dim P As TabPage
  'Set Contains = New cBoolean
  For Each P In mPageCol
    If P = Page Then Contains = True
    Exit For
  Next
End Function
Public Function IndexOf(Page As TabPage) As Long 'cInteger
Dim P As TabPage
Dim i As Long
  'Set IndexOf = New cInteger
  For Each P In mPageCol
    If P Is Page Then
      Exit Function
    End If
    i = i + 1
  Next
  IndexOf = i
End Function
Public Function Item(index As Long) As TabPage
Attribute Item.VB_UserMemId = 0
  If index > Count - 1 Then index = Count - 1
  If index < 0 Then index = 0
  If Count > 0 Then Set Item = mPageCol(index + 1)
End Function
Public Sub Remove(Page As TabPage)
Dim LngVal As Long
  LngVal = IndexOf(Page)
  Call RemoveAt(LngVal)
End Sub
Public Sub RemoveAt(index As Long)
  Call mPageCol.Remove(index + 1)
  Call SendMessage(mHandle, TCM_DELETEITEM, index, 0)
End Sub

Private Sub Class_Terminate()
  Set mPageCol = Nothing
  'werden jetzt auch alle Pages entfernt?
  'oder müßte man zuerst alle Pages durchlaufen und set nothing?
  'ausprobieren wir werden sehen
End Sub

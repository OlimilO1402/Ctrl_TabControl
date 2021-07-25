VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Implements iSubClass
'{ ====== TAB CONTROL ======================== }
Private Const WC_TABCONTROL As String = "SysTabControl"
Private Const WC_TABCONTROLA As String = "SysTabControl32"
Private Const WC_TABCONTROLW As String = "SysTabControl32"

Private Const TCS_SCROLLOPPOSITE    As Long = &H1    '// assumes multiline tab
Private Const TCS_BOTTOM            As Long = &H2
Private Const TCS_RIGHT             As Long = &H2
Private Const TCS_MULTISELECT       As Long = &H4    '// allow multi-select in button mode
Private Const TCS_FLATBUTTONS       As Long = &H8
Private Const TCS_FORCEICONLEFT     As Long = &H10
Private Const TCS_FORCELABELLEFT    As Long = &H20
Private Const TCS_HOTTRACK          As Long = &H40
Private Const TCS_VERTICAL          As Long = &H80
Private Const TCS_TABS              As Long = &H0
Private Const TCS_BUTTONS           As Long = &H100
Private Const TCS_SINGLELINE        As Long = &H0
Private Const TCS_MULTILINE         As Long = &H200
Private Const TCS_RIGHTJUSTIFY      As Long = &H0
Private Const TCS_FIXEDWIDTH        As Long = &H400
Private Const TCS_RAGGEDRIGHT       As Long = &H800
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000
Private Const TCS_OWNERDRAWFIXED    As Long = &H2000
Private Const TCS_TOOLTIPS          As Long = &H4000
Private Const TCS_FOCUSNEVER        As Long = &H8000
Private Const TCS_EX_FLATSEPARATORS As Long = &H1
Private Const TCS_EX_REGISTERDROP   As Long = &H2

Private Const TCM_FIRST             As Long = &H1300  '{ Tab control messages }   Dez: 4864

Private Const TCM_GETIMAGELIST      As Long = (TCM_FIRST + 2)
Private Const TCM_SETIMAGELIST      As Long = (TCM_FIRST + 3)
Private Const TCM_GETITEMCOUNT      As Long = (TCM_FIRST + 4)
Private Const TCM_GETITEMA          As Long = (TCM_FIRST + 5)
Private Const TCM_SETITEMA          As Long = (TCM_FIRST + 6)
Private Const TCM_INSERTITEMA       As Long = (TCM_FIRST + 7)
Private Const TCM_DELETEITEM        As Long = (TCM_FIRST + 8)
Private Const TCM_DELETEALLITEMS    As Long = (TCM_FIRST + 9)
Private Const TCM_GETITEMRECT       As Long = (TCM_FIRST + 10)
Private Const TCM_GETCURSEL         As Long = (TCM_FIRST + 11)
Private Const TCM_SETCURSEL         As Long = (TCM_FIRST + 12)
Private Const TCM_HITTEST           As Long = (TCM_FIRST + 13)
Private Const TCM_SETITEMEXTRA      As Long = (TCM_FIRST + 14)

Private Const TCM_ADJUSTRECT        As Long = (TCM_FIRST + 40)
Private Const TCM_SETITEMSIZE       As Long = (TCM_FIRST + 41)
Private Const TCM_REMOVEIMAGE       As Long = (TCM_FIRST + 42)
Private Const TCM_SETPADDING        As Long = (TCM_FIRST + 43)
Private Const TCM_GETROWCOUNT       As Long = (TCM_FIRST + 44)
Private Const TCM_GETTOOLTIPS       As Long = (TCM_FIRST + 45)
Private Const TCM_SETTOOLTIPS       As Long = (TCM_FIRST + 46)
Private Const TCM_GETCURFOCUS       As Long = (TCM_FIRST + 47)
Private Const TCM_SETCURFOCUS       As Long = (TCM_FIRST + 48)
Private Const TCM_SETMINTABWIDTH    As Long = (TCM_FIRST + 49)
Private Const TCM_DESELECTALL       As Long = (TCM_FIRST + 50)
Private Const TCM_HIGHLIGHTITEM     As Long = (TCM_FIRST + 51)
Private Const TCM_SETEXTENDEDSTYLE As Long = (TCM_FIRST + 52)
Private Const TCM_GETEXTENDEDSTYLE As Long = (TCM_FIRST + 53)

Private Const TCM_GETITEMW          As Long = (TCM_FIRST + 60)
Private Const TCM_SETITEMW          As Long = (TCM_FIRST + 61)
Private Const TCM_INSERTITEMW       As Long = (TCM_FIRST + 62)

Private Const TCM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
Private Const TCM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT

Private Const TCM_GETITEM           As Long = TCM_GETITEMA
Private Const TCM_SETITEM           As Long = TCM_SETITEMA
Private Const TCM_INSERTITEM       As Long = TCM_INSERTITEMA

Private Type InitCommonControlsEx
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

'USER32
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (ByRef TLPINITCOMMONCONTROLSEX As InitCommonControlsEx) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'GDI32
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal o As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'KERNEL32
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'die Struktur InitCommonControlsEx ist in ModTabControl.bas
Private mHandle As Long
Private mName As String           'der Name dieses Controls
Private mParent As Form           'das Formular in dem alles erstellt wird
Private WithEvents mContainer As PictureBox 'Container für mTabStrip
Attribute mContainer.VB_VarHelpID = -1
Private WithEvents mPageCont As PictureBox  'Container für alle TabPages
Attribute mPageCont.VB_VarHelpID = -1
Private mTabPages As TabPageCollection   'Liste der TabPages
Private mDelPCol As Collection    'Liste der gelöschten TabPages
Private Const DefName = "Tabelle" 'DefaultName eines Tabs (+Index)
Private mLastN As Long            'zur Erzeugung eines eindeutigen Names
Private mLastSelected As Long     'Index der zuletzt ausgewählten Seite
Private mSelectedIndex As Long    'Index der aktuell ausgewählten Seite
Private cSubClassing As cSubclass
Private mHFont As Long
Public Event Rename(sender As TabControl)
Public Event TabClick(ByVal mIndex As Long)
Public Event Paint(ByVal mIndex As Long)
'Private mLoc As New System_Drawing.Point 'Location
'Private mSize As New Size
Private mLeft As Long
Private mTop As Long
Private mWidth As Long
Private mHeight As Long
Private STPX As Long, STPY As Long

Private Sub Class_Initialize()
Dim InitCC As InitCommonControlsEx
  Set cSubClassing = New cSubclass
  InitCC.dwICC = &H8
  InitCC.dwSize = Len(InitCC)
  InitCommonControlsEx InitCC
  Set mTabPages = New TabPageCollection
  Set mDelPCol = New Collection
End Sub
Public Sub NewC(myOwner As Form, MyContainer As PictureBox, MyName As String)
Dim L As Long, T As Long, W As Long, H As Long
  mName = MyName
  Set mParent = myOwner
  If mParent.ScaleMode = 1 Then 'Twip
    STPX = Screen.TwipsPerPixelX: STPY = Screen.TwipsPerPixelY
  Else
    STPX = 1: STPY = 1
  End If
  Set mContainer = MyContainer
  mContainer.ScaleMode = mParent.ScaleMode
  'mLoc.X = mContainer.Left: mLoc.Y = mContainer.Top
  mLeft = mContainer.Left: mTop = mContainer.Top
  'mSize.Width = mContainer.ScaleWidth: mSize.Height = mContainer.ScaleHeight
  mWidth = mContainer.ScaleWidth: mHeight = mContainer.ScaleHeight
  L = 0: T = 0
  'W = mSize.Width / STPX: H = mSize.Height / STPY
  W = mWidth / STPX: H = mHeight / STPY
  mHandle = CreateWindowEx(0, WC_TABCONTROLA, 0, WS_CHILD Or WS_VISIBLE, L, T, W, H, mContainer.hwnd, 1, App.hInstance, 0)
  cSubClassing.AddMsg WM_PAINT, MSG_BEFORE
  cSubClassing.AddMsg WM_NOTIFY, MSG_BEFORE
  cSubClassing.Subclass mContainer.hwnd, Me
  mHFont = CreateFont(-MulDiv(8, GetDeviceCaps(GetDC(GetActiveWindow), LOGPIXELSY), 72), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "Microsoft Sans Serif")
  Call SendMessage(mHandle, &H30, mHFont, 1)
  If mPageCont Is Nothing Then
    Set mPageCont = mParent.Controls.Add("VB.PictureBox", mName + "mPageCont", mContainer)
    mPageCont.ScaleMode = mContainer.ScaleMode
  End If
  Call mTabPages.Create(mParent, mContainer, mPageCont, mName, mHandle)
  mContainer.BorderStyle = 0
  mContainer.Appearance = 1
End Sub
Public Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
'
End Sub
Private Sub iSubclass_Before(bHandled As Long, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
Dim mNotify As NMHDR
  Select Case uMsg
  Case WM_NOTIFY
    CopyMemory mNotify, ByVal lParam, Len(mNotify)
    If mNotify.code = TCN_SELCHANGE Then
      mLastSelected = mSelectedIndex
      mSelectedIndex = SendMessage(mHandle, TCM_GETCURSEL, 0, 0)
      RaiseEvent TabClick(mSelectedIndex)
      SelectedTab.BringToFront
      bHandled = 1
    End If
  Case WM_PAINT
    RaiseEvent Paint(mSelectedIndex)
  End Select
End Sub

Public Property Get SelectedIndex() As Long
  'mSelectedIndex = SendMessage(mHandle, TCM_GETCURSEL, 0, 0)
  SelectedIndex = mSelectedIndex
End Property
Public Property Let SelectedIndex(index As Long)
  If index < 0 Then index = 0
  mLastSelected = mSelectedIndex
  mSelectedIndex = index
  Call SendMessage(mHandle, TCM_SETCURSEL, mSelectedIndex, 0)
  RaiseEvent TabClick(mSelectedIndex)
  If Not SelectedTab Is Nothing Then SelectedTab.BringToFront
End Property

Public Property Get SelectedTab() As TabPage
  Set SelectedTab = mTabPages(mSelectedIndex)
End Property
Public Property Get TabPages() As TabPageCollection
  Set TabPages = mTabPages
End Property
Public Sub BringToFront()
  mContainer.ZOrder 0
  mPageCont.ZOrder 0
  mPageCont.Visible = True
End Sub
Public Property Get TabCount() As Long
  TabCount = mTabPages.Count
End Property
Private Sub mContainer_Resize()
Dim L As Single, T As Single, W As Single, H As Single
Dim Brdr As Long, B As Long
  L = 0:                T = 0
  W = mContainer.Width: H = mContainer.Height
  mWidth = W:  mHeight = H
  Call MoveWindow(mHandle, 0, 0, mWidth, mHeight, 1) 'mSize.Width, mSize.Height, 1)

  Brdr = 4 * STPX '4'
  B = 1 * STPX
  L = L + 1 * Brdr: T = 22 * STPX
  W = W - 2 * Brdr: H = H - T - 1 * Brdr
  mPageCont.Move L, T, W, H
  Debug.Print mContainer.Name
  'ResizeMe
End Sub
Private Sub mPageCont_Resize()
Dim i As Long, RP As TabPage
  'jetzt nur die Sichtbaren resizen
  For Each RP In mTabPages
    If RP.Visible Then RP.ResizeMe
  Next
End Sub

Public Property Get Left() As Long
  'Set Left = mLoc.X
  Left = mLeft
End Property
Public Property Let Left(LngVal As Long)
  'Set mLoc.X = LngVal
  mLeft = LngVal
End Property
Public Property Get Top() As Long
  'Set Top = mLoc.Y
  Top = mTop
End Property
Public Property Let Top(LngVal As Long)
  'Set mLoc.Y = LngVal
  mTop = LngVal
End Property
'Public Property Get Location() As System_Drawing.Point
'  Set Location = mLoc
'End Property
'Public Property Set Location(LocVal As System_Drawing.Point)
'  Set mLoc = LocVal
'End Property

Public Property Get Width() As Long
  'Set Width = mSize.Width
  Width = mWidth
End Property
Public Property Let Width(LngVal As Long)
  'Set mSize.Width = LngVal
  mWidth = LngVal
End Property
Public Property Get Height() As Long
  'Set Height = mSize.Height
  Height = mHeight
End Property
Public Property Let Height(LngVal As Long)
  'Set mSize.Height = LngVal
  mHeight = LngVal
End Property
'Public Property Get Size() As Size
'  Set Size = mSize
'End Property
'Public Property Set Size(SzVal As Size)
'  Set mSize = SzVal
'  Call ResizeMe
'End Property

'Public Sub ResizeMe()
  'mContainer_Resize
  'Call mContainer.Move(mContainer.Left, mContainer.Top, mWidth, mHeight) 'mSize.Width, mSize.Height)
  'Call MoveWindow(mHandle, 0, 0, mWidth, mHeight, 1) 'mSize.Width, mSize.Height, 1)
  'Call mPageCont.Move
'End Sub

Private Sub Class_Terminate()
  Set mContainer = Nothing  'Container für mTabStrip
  Set mPageCont = Nothing   'Container für alle TabPages
  Set mTabPages = Nothing   'Liste der TabPages
  If mHFont <> 0 Then Call DeleteObject(mHFont)
  If mHandle <> 0 Then Call DestroyWindow(mHandle)
End Sub

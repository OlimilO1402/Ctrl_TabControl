Attribute VB_Name = "ModTabControl"
Option Explicit
Public Const LVM_FIRST            As Long = &H1000  '{ ListView messages }      Dez: 4096
Public Const TV_FIRST             As Long = &H1100  '{ TreeView messages }      Dez: 4352
Public Const HDM_FIRST            As Long = &H1200  '{ Header messages }        Dez: 4608
Public Const TCM_FIRST            As Long = &H1300  '{ Tab control messages }   Dez: 4864
Public Const PGM_FIRST            As Long = &H1400  '{ Pager control messages } Dez: 5120
Public Const CCM_FIRST            As Long = &H2000  '{ Common control shared messages } 8192

Public Const CCM_SETBKCOLOR       As Long = CCM_FIRST + 1 ' // lParam is bkColor
Public Const CCM_SETCOLORSCHEME   As Long = CCM_FIRST + 2 ' // lParam is color scheme
Public Const CCM_GETCOLORSCHEME   As Long = CCM_FIRST + 3 ' // fills in COLORSCHEME pointed to by lParam
Public Const CCM_GETDROPTARGET    As Long = CCM_FIRST + 4
Public Const CCM_SETUNICODEFORMAT As Long = CCM_FIRST + 5
Public Const CCM_GETUNICODEFORMAT As Long = CCM_FIRST + 6


'{ ====== WM_NOTIFY codes (NMHDR.code values) ================== }
Public Const NM_FIRST             As Long = 0      ' { generic to all controls }
Public Const NM_LAST              As Long = (-99)

Public Const TCN_FIRST            As Long = -550
Public Const TCN_LAST             As Long = (-580)
Public Const TCN_KEYDOWN          As Long = (TCN_FIRST - 0)
Public Const TCN_SELCHANGE        As Long = (TCN_FIRST - 1)
Public Const TCN_SELCHANGING      As Long = (TCN_FIRST - 2)
Public Const TCN_GETOBJECT        As Long = (TCN_FIRST - 3)
Public Const TCN_FOCUSCHANGE      As Long = (TCN_FIRST - 4)

'Mouse HitTest
Public Const TCHT_NOWHERE         As Long = &H1
Public Const TCHT_ONITEMICON      As Long = &H2
Public Const TCHT_ONITEMLABEL     As Long = &H4
Public Const TCHT_ONITEM = TCHT_ONITEMICON Or TCHT_ONITEMLABEL

'{ ====== COMMON CONTROL STYLES ================ }

Public Const CCS_TOP              As Long = &H1
Public Const CCS_NOMOVEY          As Long = &H2
Public Const CCS_BOTTOM           As Long = &H3
Public Const CCS_NORESIZE         As Long = &H4
Public Const CCS_NOPARENTALIGN    As Long = &H8
Public Const CCS_ADJUSTABLE       As Long = &H20
Public Const CCS_NODIVIDER        As Long = &H40
Public Const CCS_VERT             As Long = &H80
Public Const CCS_LEFT             As Long = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT            As Long = (CCS_VERT Or CCS_BOTTOM)
Public Const CCS_NOMOVEX          As Long = (CCS_VERT Or CCS_NOMOVEY)


'{ Window Styles }
Public Const WS_OVERLAPPED        As Long = &H0&
Public Const WS_POPUP             As Long = &H80000000
Public Const WS_CHILD             As Long = &H40000000
Public Const WS_MINIMIZE          As Long = &H20000000
Public Const WS_VISIBLE           As Long = &H10000000
Public Const WS_DISABLED          As Long = &H8000000
Public Const WS_CLIPSIBLINGS      As Long = &H4000000
Public Const WS_CLIPCHILDREN      As Long = &H2000000
Public Const WS_MAXIMIZE          As Long = &H1000000
Public Const WS_CAPTION           As Long = &HC00000    ' { WS_BORDER or WS_DLGFRAME  }
Public Const WS_BORDER            As Long = &H800000
Public Const WS_DLGFRAME          As Long = &H400000
Public Const WS_VSCROLL           As Long = &H200000
Public Const WS_HSCROLL           As Long = &H100000
Public Const WS_SYSMENU           As Long = &H80000
Public Const WS_THICKFRAME        As Long = &H40000
Public Const WS_GROUP             As Long = &H20000
Public Const WS_TABSTOP           As Long = &H10000

Public Const WS_MINIMIZEBOX       As Long = &H20000
Public Const WS_MAXIMIZEBOX       As Long = &H10000

Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_SIZEBOX = WS_THICKFRAME

'{ Common Window Styles }
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_CHILDWINDOW = (WS_CHILD)

'{ Extended Window Styles }
Public Const WS_EX_DLGMODALFRAME  As Long = 1
Public Const WS_EX_NOPARENTNOTIFY As Long = 4
Public Const WS_EX_TOPMOST        As Long = 8
Public Const WS_EX_ACCEPTFILES    As Long = &H10
Public Const WS_EX_TRANSPARENT    As Long = &H20
Public Const WS_EX_MDICHILD       As Long = &H40
Public Const WS_EX_TOOLWINDOW     As Long = &H80
Public Const WS_EX_WINDOWEDGE     As Long = &H100
Public Const WS_EX_CLIENTEDGE     As Long = &H200
Public Const WS_EX_CONTEXTHELP    As Long = &H400

Public Const WS_EX_RIGHT          As Long = &H1000
Public Const WS_EX_LEFT           As Long = &H0&
Public Const WS_EX_RTLREADING     As Long = &H2000
Public Const WS_EX_LTRREADING     As Long = &H0&
Public Const WS_EX_LEFTSCROLLBAR  As Long = &H4000
Public Const WS_EX_RIGHTSCROLLBAR As Long = &H0&

Public Const WS_EX_CONTROLPARENT  As Long = &H10000
Public Const WS_EX_STATICEDGE     As Long = &H20000
Public Const WS_EX_APPWINDOW      As Long = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW  As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

Public Const WS_EX_LAYERED        As Long = &H80000
Public Const WS_EX_NOINHERITLAYOUT As Long = &H100000     ' // Disable inheritence of mirroring by children
Public Const WS_EX_LAYOUTRTL      As Long = &H400000   ' // Right to left mirroring
Public Const WS_EX_COMPOSITED     As Long = &H2000000
Public Const WS_EX_NOACTIVATE     As Long = &H8000000

Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90

Public Type InitCommonControlsEx
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type

Public Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Public Type TCITEM 'TabControlItem = Tabber
    mask As Long
    dwState As Long
    dwStateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


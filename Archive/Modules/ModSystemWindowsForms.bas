Attribute VB_Name = "ModSystemWindowsForms"
Option Explicit
'eigentlich auch eine Klasse
'Public Enum DialogResult 'System.Windows.Forms.DialogResult
'  None = 0
'  OK = 1
'  Cancel = 2
'  Abort = 3
'  Retry = 4
'  Ignore = 5
'  Yes = 6
'  No = 7
'End Enum

'Public Enum E_DockStyle
'  DockStyleNone = 0
'  DockStyleTop = 1
'  DockStyleBottom = 2
'  DockStyleLeft = 3
'  DockStyleRight = 4
'  DockStyleFill = 5
'End Enum

Public Enum FormBorderStyle
  None = 0
  FixedSingle = 1
  Fixed3D = 2
  FixedDialog = 3
  Sizable = 4
  FixedToolWindow = 5
  SizableToolWindow = 6
End Enum

Public Enum FormStartPosition
  Manual = 0
  CenterScreen = 1
  WindowsDefaultLocation = 2
  WindowsDefaultBounds = 3
  CenterParent = 4
End Enum

Public Enum FormWindowState
  Normal = 0
  Minimized = 1
  Maximized = 2
End Enum

Public Enum FrameStyle
  Dashed = 0
  Thick = 1
End Enum

Public Enum MdiLayout
  Cascade = 0
  TileHorizontal = 1
  TileVertical = 2
  ArrangeIcons = 3
End Enum

Public Enum MouseButtons
   MouseButtonsNone = 0
   MouseButtonsLeft = &H100000       '1048576
   MouseButtonsRight = &H200000      '2097152
   MouseButtonsMiddle = &H400000     '4194304
   MouseButtonsXButton1 = &H800000   '8388608
   MouseButtonsXButton2 = &H1000000 '16777216
End Enum

Public Enum Keys
  KeysNone = 0
  KeysLButton = 1
  KeysRButton = 2
  KeysCancel = 3
  KeysMButton = 4
  KeysXButton1 = 5
  KeysXButton2 = 6
  
  KeysBack = 8
  KeysTab = 9
  KeysLineFeed = 10
  
  KeysClear = 12
  KeysEnter = 13
  KeysReturn = 13
  
  KeysShiftKey = 16   'Die UMSCHALTTASTE.
  KeysControlKey = 17
  KeysMenu = 18
  KeysPause = 19
  KeysCapital = 20
  KeysCapsLock = 20
  KeysHanguelMode = 21
  KeysHangulMode = 21
  KeysKanaMode = 21

  KeysJunjaMode = 23
  KeysFinalMode = 24
  KeysHanjaMode = 25
  KeysKanjiMode = 25

  KeysEscape = 27
  KeysIMEConvert = 28
  KeysIMENonconvert = 29
  KeysIMEAceept = 30
  KeysIMEModeChange = 31
  KeysSpace = 32    'Die LEERTASTE.
  KeysPageUp = 33   'Die BILD-AUF-TASTE.
  KeysPrior = 33    'Die BILD-AUF-TASTE.
  KeysPageDown = 34 'Die BILD-AB-TASTE.
  KeysNext = 34     'Die BILD-AB-TASTE.
  KeysEnd = 35      'Die ENDE-TASTE.
  KeysHome = 36     'Die POS1-TASTE.
  KeysLeft = 37
  KeysUp = 38
  KeysRight = 39
  KeysDown = 40
  KeysSelect = 41
  KeysPrint = 42
  KeysExecute = 43
  KeysSnapshot = 44    'Die DRUCK-TASTE.
  KeysPrintScreen = 44 'Die DRUCK-TASTE.
  KeysInsert = 45
  KeysDelete = 46
  KeysHelp = 47
  KeysD0 = 48
  KeysD1 = 49
  KeysD2 = 50
  KeysD3 = 51
  KeysD4 = 52
  KeysD5 = 53
  KeysD6 = 54
  KeysD7 = 55
  KeysD8 = 56
  KeysD9 = 57

  KeysA = 65
  KeysB = 66
  KeysC = 67
  KeysD = 68
  KeysE = 69
  KeysF = 70
  KeysG = 71
  KeysH = 72
  KeysI = 73
  KeysJ = 74
  KeysK = 75
  KeysL = 76
  KeysM = 77
  KeysN = 78
  KeysO = 79
  KeysP = 80
  KeysQ = 81
  KeysR = 82
  KeysS = 83
  KeysT = 84
  KeysU = 85
  KeysV = 86
  KeysW = 87
  KeysX = 88
  KeysY = 89
  KeysZ = 90
  KeysLWin = 91 'Die linke WINDOWS-TASTE (Microsoft Natural Keyboard).
  KeysRWin = 92 'Die rechte WINDOWS-TASTE (Microsoft Natural Keyboard).
  KeysApps = 93 'Die ANWENDUNGSTASTE (Microsoft Natural Keyboard).
  
  KeysNumPad0 = 96   'Die 0-TASTE auf der Zehnertastatur.
  KeysNumPad1 = 97   'Die 1-TASTE auf der Zehnertastatur.
  KeysNumPad2 = 98   'Die 2-TASTE auf der Zehnertastatur.
  KeysNumPad3 = 99   'Die 3-TASTE auf der Zehnertastatur.
  KeysNumPad4 = 100  'Die 4-TASTE auf der Zehnertastatur.
  KeysNumPad5 = 101  'Die 5-TASTE auf der Zehnertastatur.
  KeysNumPad6 = 102  'Die 6-TASTE auf der Zehnertastatur.
  KeysNumPad7 = 103  'Die 7-TASTE auf der Zehnertastatur.
  KeysNumPad8 = 104  'Die 8-TASTE auf der Zehnertastatur.
  KeysNumPad9 = 105  'Die 9-TASTE auf der Zehnertastatur.
  KeysMultiply = 106 'Die MULTIPLIKATIONSTASTE.
  KeysAdd = 107      'Die ADDITIONSTASTE.
  KeysSeparator = 108 'Die TRENNZEICHENTASTE.
  KeysSubtract = 109  'Die SUBTRAKTIONSTASTE.
  KeysDecimal = 110  ' Die KOMMATASTE auf der Zehnertastatur.
  KeysDivide = 111
  KeysF1 = 112
  KeysF2 = 113
  KeysF3 = 114
  KeysF4 = 115
  KeysF5 = 116
  KeysF6 = 117
  KeysF7 = 118
  KeysF8 = 119
  KeysF9 = 120
  KeysF10 = 121
  KeysF11 = 122
  KeysF12 = 123
  KeysF13 = 124
  KeysF14 = 125
  KeysF15 = 126
  KeysF16 = 127
  KeysF17 = 128
  KeysF18 = 129
  KeysF19 = 130
  KeysF20 = 131
  KeysF21 = 132
  KeysF22 = 133
  KeysF23 = 134
  KeysF24 = 135
  KeysNumLock = 144
  KeysScroll = 145

  KeysLShiftKey = 160
  KeysRShiftKey = 161
  KeysLControlKey = 162
  KeysRControlKey = 163
  KeysLMenu = 164
  KeysRMenu = 165
  BrowserBack = 166      'Die BROWSER-ZURÜCK-TASTE (Windows 2000 oder höher).
  BrowserForward = 167   'Die BROWSER-VORWÄRTS-TASTE (Windows 2000 oder höher).
  BrowserRefresh = 168   'Die BROWSER-AKTUALISIEREN-TASTE (Windows 2000 oder höher).
  BrowserStop = 169      'Die BROWSER-ABBRECHEN-TASTE (Windows 2000 oder höher).
  BrowserSearch = 170    'Die BROWSER-SUCHEN-TASTE (Windows 2000 oder höher).
  BrowserFavorites = 171 'Die BROWSER-FAVORITEN-TASTE (Windows 2000 oder höher).
  BrowserHome = 172      'Die BROWSER-STARTSEITE-TASTE (Windows 2000 oder höher).
  KeysVolumeMute = 173   'Die Taste zum Stummschalten (Windows 2000 oder höher).
  KeysVolumeDown = 174   'Die Taste zum Verringern der Lautstärke (Windows 2000 oder höher).
  KeysVolumeUp = 175     'Die Taste zum Erhöhen der Lautstärke (Windows 2000 oder höher).
  KeysMediaNextTrack = 176
  KeysMediaPreviousTrack = 177
  KeysMediaStop = 178
  KeysMediaPlayPause = 179
  KeysLaunchMail = 180
  KeysSelectMedia = 181
  KeysLaunchApplication1 = 182
  KeysLaunchApplication2 = 183

  KeysOemSemicolon = 186
  KeysOemplus = 187
  KeysOemcomma = 188
  KeysOemMinus = 189
  KeysOemPeriod = 190
  KeysOemQuestion = 191
  KeysOemtilde = 192
  KeysOemOpenBrackets = 219
  KeysOemPipe = 220
  KeysOemCloseBrackets = 221
  KeysOemQuotes = 222
  KeysOem8 = 223
  KeysOemBackslash = 226
  
  KeysProcessKey = 229

  KeysAttn = 246         'Die ATTN-TASTE.
  KeysCrsel = 247        'Die CRSEL-TASTE.
  KeysExsel = 248        'Die EXSEL-TASTE.
  KeysEraseEof = 249     'Die ERASE EOF-TASTE.
  KeysPlay = 250         'Die PLAY-TASTE.
  KeysZoom = 251         'Die ZOOM-TASTE.
  KeysNoName = 252       'Eine für die zukünftige Verwendung reservierte Konstante
  KeysPa1 = 253          'Die PA1-TASTE. '(OM: vielleicht Home? bzw. Pos1?)
  KeysOemClear = 254     'Die CLEAR-TASTE.
  
  KeysKeyCode = 65535    'Die Bitmaske zum Extrahieren eines Tastencodes aus einem Tastenwert.
  KeysModifiers = -65536 'Die Bitmaske zum Extrahieren von Modifizierern aus einem Tastenwert.
  KeysShift = 65536      'Die Modifizierertaste UMSCHALT.
  KeysControl = 131072   'Die Modifizierertaste STRG.
  KeysAlt = 262144       'Die Modifizierertaste ALT.
End Enum

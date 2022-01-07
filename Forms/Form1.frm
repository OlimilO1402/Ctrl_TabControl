VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   5460
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PnlTabPage5 
      BackColor       =   &H80000005&
      Height          =   3735
      Left            =   7440
      ScaleHeight     =   3675
      ScaleWidth      =   5835
      TabIndex        =   33
      Top             =   1560
      Width           =   5895
   End
   Begin VB.PictureBox PnlTabPage4 
      BackColor       =   &H80000005&
      Height          =   3735
      Left            =   7200
      ScaleHeight     =   3675
      ScaleWidth      =   5835
      TabIndex        =   32
      Top             =   1320
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4920
      TabIndex        =   31
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox PnlStatusBar 
      Align           =   2  'Unten ausrichten
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   14520
      TabIndex        =   30
      Top             =   5085
      Width           =   14520
   End
   Begin VB.PictureBox PnlTabCtrl2 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   6960
      ScaleHeight     =   3915
      ScaleWidth      =   5835
      TabIndex        =   29
      Top             =   1080
      Width           =   5895
   End
   Begin VB.PictureBox PnlTabPage2 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   6600
      ScaleHeight     =   3915
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   840
      Width           =   6015
      Begin VB.PictureBox Picture2 
         Height          =   1935
         Left            =   3720
         ScaleHeight     =   1875
         ScaleWidth      =   1995
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   3375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command7"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command7"
         Height          =   375
         Left            =   3720
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check2"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1095
      End
   End
   Begin VB.PictureBox PnlTabPage1 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   6360
      ScaleHeight     =   3915
      ScaleWidth      =   5955
      TabIndex        =   16
      Top             =   600
      Width           =   6015
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4680
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1635
         ScaleWidth      =   5475
         TabIndex        =   17
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.CommandButton BtnCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton BtnMove 
      Caption         =   "Move"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton BtnRename 
      Caption         =   "Rename"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton BtnDel 
      Caption         =   "Del -"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton BtnAdd 
      Caption         =   "Add +"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox PnlTabCtrl1 
      BackColor       =   &H80000005&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4515
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   480
      Width           =   6045
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_DefTabName As String
Private m_LastNr     As Long
Private WithEvents TabControl1 As TabControl
Attribute TabControl1.VB_VarHelpID = -1
Private WithEvents TabControl2 As TabControl
Attribute TabControl2.VB_VarHelpID = -1
Private m_CurTabCtrl As TabControl

Private Sub Command1_Click()
    MsgBox Me.PnlTabCtrl1.Width & " " & Me.PnlTabCtrl1.Height
    MsgBox Me.PnlTabCtrl2.Width & " " & Me.PnlTabCtrl2.Height
End Sub

Private Sub Form_Load()
    
    m_DefTabName = "Tabelle"
    
    Me.Caption = App.EXEName
    Me.Width = 6345
    
    'Erstes unterstes TabControl
    Set TabControl1 = MNew.TabControl(Me, PnlTabCtrl1, "TabControl1")
    
    Set m_CurTabCtrl = TabControl1
    
    NewTabPage TabControl1, "TabPage1", Me.PnlTabPage1
    
    NewTabPage TabControl1, "TabPage2", Me.PnlTabPage2
        
    Dim bkColor1 As Long: bkColor1 = GetBkColor(GetDC(PnlTabPage1.hwnd))
    BackgroundColorAndAllChildren(PnlTabPage1, Nothing) = bkColor1
    
    Dim bkColor2 As Long: bkColor2 = GetBkColor(GetDC(PnlTabPage2.hwnd))
    BackgroundColorAndAllChildren(PnlTabPage2, Nothing) = bkColor2
    
    NewTabPage TabControl1, "TabPage3", Me.PnlTabCtrl2
    
    'Zweites oberstes TabControl
    Set TabControl2 = MNew.TabControl(Me, PnlTabCtrl2, "TabControl2")
    
    NewTabPage TabControl2, "TabPage4", Me.PnlTabPage4
    
    NewTabPage TabControl2, "TabPage5", Me.PnlTabPage5
    
    Dim bkColor4 As Long: bkColor4 = GetBkColor(GetDC(PnlTabPage4.hwnd))
    BackgroundColorAndAllChildren(PnlTabPage4, Nothing) = bkColor4
    
    Dim bkColor5 As Long: bkColor5 = GetBkColor(GetDC(PnlTabPage5.hwnd))
    BackgroundColorAndAllChildren(PnlTabPage5, Nothing) = bkColor5
    
End Sub

'TODO OM: collection of controls to exclude
'e.g. if you want only the controls on the Form colored and not on the Picturebox
Public Property Let BackgroundColorAndAllChildren(Ctrl, CtrlsToExclude As Collection, ByVal Color As Long)
Try: On Error GoTo Catch
    Ctrl.BackColor = Color
    Debug.Print TypeName(Ctrl) & " : " & Ctrl.Name
    Dim C
    For Each C In Me.Controls
        If C.Container Is Ctrl Then
            If Not C.Container Is Nothing Then
                BackgroundColorAndAllChildren(C, Nothing) = Color
            End If
        End If
    Next
    Exit Property
Catch: On Error GoTo 0
End Property

Private Function NewTabPage(TC As TabControl, Name As String, Optional Ctrl As PictureBox = Nothing) As TabPage
'Achtung Reihenfolge beachten:
'zuerst
' * TabPages.add(NewTabPage),
'dann
' * NewTabPage.Controls.Add(Ctrl)
    Set NewTabPage = New TabPage: NewTabPage.Text = Name
    Call TC.TabPages.Add(NewTabPage)
    If Not Ctrl Is Nothing Then
        Ctrl.BorderStyle = 0
        Call NewTabPage.Controls.Add(Ctrl)
    End If
End Function

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub TabControl1_Paint(ByVal mIndex As Long)
    '
End Sub

Private Sub TabControl1_Rename(sender As TabControl)
    '
End Sub

'Private Sub TabControl1_TabClick(ByVal mIndex As Long)
'    Select Case mIndex
'    Case 0: tp1.BringToFront
'            TabControl1.TabPages.Item(1).BringToFront
'    Case 1: tp2.BringToFront
'            TabControl1.TabPages.Item(1).BringToFront
'    End Select
'End Sub

Private Sub BtnAdd_Click()
    Dim s As String: s = GetNewUniqueName
    s = InputBox("Geben Sie bitte einen Tabellennamen an: ", "Tabellenname?", s, , , 0, 0)
    If StrPtr(s) = 0 Then Exit Sub
    Dim NewPage As TabPage: Set NewPage = NewTabPage(m_CurTabCtrl, s, Nothing)
    'NewPage.Text = s
    'm_CurTabCtrl.TabPages.Add NewPage
    'm_CurTabCtrl.SelectedIndex = m_CurTabCtrl.TabCount - 1
    m_LastNr = m_LastNr + 1
End Sub

Private Sub BtnDel_Click()
    '
End Sub

Private Sub BtnRename_Click()
    '
End Sub

Private Sub BtnMove_Click()
    '
End Sub

Private Sub BtnCopy_Click()
    '
End Sub


Private Function GetNewUniqueName() As String
    Dim n As Long: n = m_LastNr + 1
    Dim NUniName As String: NUniName = m_DefTabName + CStr(n)
    
    If NameExists(NUniName) Then
        Dim i As Long
        For i = 0 To m_CurTabCtrl.TabPages.Count - 1
            NUniName = m_DefTabName + CStr(i)
            If Not NameExists(NUniName) Then Exit For
        Next
    End If
    GetNewUniqueName = NUniName
End Function

Private Function NameExists(StrNam As String) As Boolean
    Dim p As TabPage
    'Alle TabControls durchsuchen!
    If TabControl1 Is Nothing Then Exit Function
    For Each p In TabControl1.TabPages
        If p.Text = StrNam Then
            NameExists = True
            Exit Function
        End If
    Next
    If TabControl2 Is Nothing Then Exit Function
    For Each p In TabControl2.TabPages
        If p.Text = StrNam Then
            NameExists = True
            Exit Function
        End If
    Next
End Function

Private Function GetTabPageByName(Value As String) As TabPage
    For Each GetTabPageByName In m_CurTabCtrl.TabPages
        If GetTabPageByName.Text = Value Then Exit Function
    Next
End Function

Private Sub TabControl1_TabClick(ByVal Index As Long)
    Set m_CurTabCtrl = TabControl1
    m_LastNr = Index
End Sub
Private Sub TabControl2_TabClick(ByVal Index As Long)
    Set m_CurTabCtrl = TabControl2
    m_LastNr = Index
End Sub


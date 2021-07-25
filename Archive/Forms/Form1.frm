VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PicOwnStatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6315
      TabIndex        =   24
      Top             =   3960
      Width           =   6375
   End
   Begin VB.PictureBox Panel3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3120
      ScaleHeight     =   1215
      ScaleWidth      =   3135
      TabIndex        =   15
      Top             =   2160
      Width           =   3135
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   18
         Text            =   "TextBox3"
         Top             =   480
         Width           =   1080
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Text            =   "TextBox2"
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Text            =   "TextBox1"
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.PictureBox Panel4 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   2880
      ScaleHeight     =   1215
      ScaleWidth      =   3135
      TabIndex        =   19
      Top             =   1560
      Width           =   3135
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FEFCFC&
         Caption         =   "CheckBox1"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   1320
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FEFCFC&
         Caption         =   "CheckBox2"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   120
         Width           =   1320
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FEFCFC&
         Caption         =   "RadioButton1"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   480
         Width           =   1320
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FEFCFC&
         Caption         =   "RadioButton2"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   840
         Width           =   1320
      End
   End
   Begin VB.CommandButton BtnCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnMove 
      Caption         =   "Move"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnRen 
      Caption         =   "Rename"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnDel 
      Caption         =   "Del"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton BtnAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox Panel2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3000
      ScaleHeight     =   855
      ScaleWidth      =   3135
      TabIndex        =   7
      Top             =   720
      Width           =   3135
      Begin VB.OptionButton RadioButton2 
         BackColor       =   &H00FEFCFC&
         Caption         =   "RadioButton2"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   1320
      End
      Begin VB.OptionButton RadioButton1 
         BackColor       =   &H00FEFCFC&
         Caption         =   "RadioButton1"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   120
         Width           =   1320
      End
      Begin VB.CheckBox CheckBox2 
         BackColor       =   &H00FEFCFC&
         Caption         =   "CheckBox2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1320
      End
      Begin VB.CheckBox CheckBox1 
         BackColor       =   &H00FEFCFC&
         Caption         =   "CheckBox1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.PictureBox Panel1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1320
      ScaleHeight     =   1455
      ScaleWidth      =   1575
      TabIndex        =   6
      Top             =   720
      Width           =   1575
      Begin VB.TextBox TextBox1 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "TextBox1"
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox TextBox2 
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Text            =   "TextBox2"
         Top             =   480
         Width           =   1080
      End
      Begin VB.TextBox TextBox3 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Text            =   "TextBox3"
         Top             =   840
         Width           =   1080
      End
   End
   Begin VB.PictureBox PnlPicBTabC1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      ScaleHeight     =   161
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mTabControl1 As TabControl
Attribute mTabControl1.VB_VarHelpID = -1
Private WithEvents mTabControl2 As TabControl
Attribute mTabControl2.VB_VarHelpID = -1
Private mLastN As Long 'New cInteger
Private mDefTabName As String
Private mCurTabC As TabControl

Private Sub Form_Load()
Dim NewPage As TabPage
  mLastN = 0
  mDefTabName = "Tabelle"

  'Me.Move 0, 0
  Set mTabControl1 = New_TabControl(Me, PnlPicBTabC1, "mTabControl1")

  Set NewPage = New TabPage
  NewPage.Text = mDefTabName & CStr(mTabControl1.TabPages.Count + 1)
  Call mTabControl1.TabPages.Add(NewPage)
  Call NewPage.Controls.Add(Panel1)

  Set NewPage = New TabPage
  NewPage.Text = mDefTabName & CStr(mTabControl1.TabPages.Count + 1)
  Call mTabControl1.TabPages.Add(NewPage)
  Call NewPage.Controls.Add(Panel2)

  Set NewPage = New TabPage
  NewPage.Text = "Tabelle3"
  Call mTabControl1.TabPages.Add(NewPage)
  

  Set mTabControl2 = New_TabControl(Me, NewPage.Page, "mTabControl2")

  Set NewPage = New TabPage
  NewPage.Text = mDefTabName & CStr(mTabControl2.TabPages.Count + 1)
  Call mTabControl2.TabPages.Add(NewPage)
  Call NewPage.Controls.Add(Panel3)

  Set NewPage = New TabPage
  NewPage.Text = mDefTabName & CStr(mTabControl2.TabPages.Count + 1)
  Call mTabControl2.TabPages.Add(NewPage)
  Call NewPage.Controls.Add(Panel4)
  
  mTabControl1.SelectedIndex = 0
  mLastN = mTabControl1.TabCount
  mTabControl2.SelectedIndex = 0
End Sub

Private Sub BtnAdd_Click()
Dim StrNam As String, NewPage As New TabPage
  StrNam = GetNewUniqueName
  StrNam = InputBox("Geben Sie bitte einen Tabellennamen an: ", "Tabellenname?", StrNam, , , 0, 0)
  If StrNam <> "" Then
    NewPage.Text = StrNam
    Call mCurTabC.TabPages.Add(NewPage)
    mCurTabC.SelectedIndex = mCurTabC.TabCount - 1
    mLastN = mLastN + 1
  End If
End Sub

Private Sub BtnDel_Click()
  Dim mR As VbMsgBoxResult, DelPage As New TabPage, DelText As String, SelI As Long 'New cInteger
  If mCurTabC.TabCount > 0 Then
    SelI = mCurTabC.SelectedIndex
    DelText = mCurTabC.SelectedTab.Text
    mR = MsgBox("Möchten Sie den  aktuellen Tab löschen: " & """" & DelText & """" & vbCrLf & "(Nein, d.h.: Geben Sie den Tabellennamen an)", vbYesNoCancel, "Bitte Löschen bestätigen!", 0, 0)
    Select Case mR
    Case vbYes
      Call mCurTabC.TabPages.Remove(mCurTabC.SelectedTab)
      mCurTabC.SelectedIndex = SelI - 1
    Case vbNo
      DelText = InputBox("Geben Sie den Namen der zu löschenden Tabelle an:", "Tabellenname?", DelText, , , 0, 0)
      If DelText <> "" Then
        'jetzt zuerst rausfinden  ob es den angegebenen Tab überhaupt gibt
        If Not GetTabPageByName(DelText) Is Nothing Then
          Call mCurTabC.TabPages.Remove(GetTabPageByName(DelText))
          mCurTabC.SelectedIndex = SelI - 1
        End If
      End If
    End Select
  Else
    MsgBox "TabControl enthält noch keine Tabs"
  End If
End Sub

Private Sub BtnRen_Click()
Dim NewText As String
  If mCurTabC.TabCount > 0 Then
    NewText = mCurTabC.TabPages(mCurTabC.SelectedIndex).Text
    NewText = InputBox("Geben sie den neuen Namen der aktuellen Tabelle an: ", "Tabellenname?", NewText, , , 0, 0)
    If NewText <> "" Then
      mCurTabC.TabPages(mCurTabC.SelectedIndex).Text = NewText
    End If
  Else
    MsgBox "TabControl enthält noch keine Tabs"
  End If
End Sub


Private Sub BtnMove_Click()
Dim CurPos As Long, NewPos As Long, StrNewPos As String
Dim ThisTabPage As TabPage
  If mCurTabC.TabCount > 0 Then
    CurPos = mCurTabC.SelectedIndex
    StrNewPos = InputBox("An welche Position möchten sie die Tabelle verschieben, geben Sie bitte einen Index an: ", "Tabellenposition?", CStr(CurPos), , , 0, 0)
    If StrNewPos <> "" And IsNumeric(StrNewPos) Then NewPos = CLng(StrNewPos)
    Set ThisTabPage = mCurTabC.TabPages.Item(CurPos)
    'remove ThisTabPage at the current position
    
    'Add ThisTabPage at the new position
    
  End If
End Sub
Private Sub BtnCopy_Click()
  '
End Sub

Private Sub Form_Resize()
Dim L As Long, T As Long, W As Long, H As Long
Dim Brdr As Long ' New cInteger
'Dim mSz As New ClsSize
  'Me.ScaleMode = vbPixels
  Brdr = 8 '* Screen.TwipsPerPixelX
  'mSz.Width = Me.ScaleWidth - 2 * Brdr
  'mSz.Height = Me.ScaleHeight - mTabControl1.Top - 1 * Brdr
  'Set mTabControl1. = mSz
  L = PnlPicBTabC1.Left: T = PnlPicBTabC1.Top
  W = Me.ScaleWidth - 2 * Brdr
  H = Me.ScaleHeight - T - PicOwnStatusBar.Height - 1 * Brdr
  PnlPicBTabC1.Move L, T, W, H
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Application.eExit
End Sub

Private Sub mTabControl1_Rename(sender As TabControl) 'wird ausgelöst bei Tip auf F2
  Call BtnRen_Click
End Sub

Private Function GetNewUniqueName() As String 'ambigeous ist schmarrn das heißt genau das gegenteil
'Dim n As New cInteger,
Dim n As Long, i As Long, NUniName As String
  n = mLastN + 1
  NUniName = mDefTabName + CStr(n) '.ToString
  If NameExists(NUniName) Then
    For i = 0 To mCurTabC.TabPages.Count - 1
      NUniName = mDefTabName + CStr(i)
      If Not NameExists(NUniName) Then Exit For
    Next
  End If
  GetNewUniqueName = NUniName
End Function

Private Function NameExists(StrNam As String) As Boolean
Dim P As TabPage
  'Alle TabControls durchsuchen!
  For Each P In mTabControl1.TabPages
    If P.Text = StrNam Then
      NameExists = True
      Exit Function
    End If
  Next
  For Each P In mTabControl2.TabPages
    If P.Text = StrNam Then
      NameExists = True
      Exit Function
    End If
  Next
End Function

Private Function GetTabPageByName(StrVal As String) As TabPage
  For Each GetTabPageByName In mCurTabC.TabPages
    If GetTabPageByName.Text = StrVal Then Exit Function
  Next
End Function

Private Sub mTabControl1_TabClick(ByVal mIndex As Long)
  Set mCurTabC = mTabControl1
  mLastN = mIndex
End Sub
Private Sub mTabControl2_TabClick(ByVal mIndex As Long)
  Set mCurTabC = mTabControl2
  mLastN = mIndex
End Sub


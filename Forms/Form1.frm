VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   6255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PnlTabPage2 
      BackColor       =   &H80000005&
      Height          =   3975
      Left            =   6840
      ScaleHeight     =   3915
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   960
      Width           =   6015
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000005&
         Height          =   1935
         Left            =   3720
         ScaleHeight     =   1875
         ScaleWidth      =   1995
         TabIndex        =   15
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
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
         BackColor       =   &H80000005&
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "Option1"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000005&
         Caption         =   "Check2"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000005&
         Caption         =   "Check2"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
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
         BackColor       =   &H80000005&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1635
         ScaleWidth      =   5475
         TabIndex        =   17
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox PnlTabCtrl 
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents TabControl1 As TabControl
Attribute TabControl1.VB_VarHelpID = -1

Private Sub Form_Load()
    
    Set TabControl1 = MNew.TabControl(Me, PnlTabCtrl, "TabControl1")
    
    NewTabPage TabControl1, "Tab1", Me.PnlTabPage1
    
    NewTabPage TabControl1, "Tab2", Me.PnlTabPage2
    
End Sub

'Private Sub InitTabStrips()
'
'    Set TabStrip1 = New_TabControl(Me, PicOwnTab1, "TabStrip1")
'  '###############  Tabpages init ##################
'    Call NewTabPage(TabStrip1, "Grundwerte", FraBasicVal)
'    Call NewTabPage(TabStrip1, "WALLS", FraAdWLS)
'    Call NewTabPage(TabStrip1, "KEM/Gleitkreis", FraAdKEM)
'
'    Dim TC1Page4 As TabPage: Set TC1Page4 = NewTabPage(TabStrip1, "FEM")
'
'  '##########  das zweite TabControl
'    Set TabStrip2 = New_TabControl(Me, TC1Page4.Page, "TabStrip2")
'
'    Call NewTabPage(TabStrip2, "Elast/Plast", FraAdFEMallg)
'    Call NewTabPage(TabStrip2, "Wand-Boden", FraAdFEMwabo)
'
'    '2009_12_23 OM: das Sofistik Materialgesetz "Duncan Chang" gibts nicht mehr
'    'Call NewTabPage(TabStrip2, "Duncan Chang", FraAdFEMdunc)
'
'    Call NewTabPage(TabStrip2, "GRAN", FraAdFEMgran)
'
'End Sub
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

VERSION 5.00
Begin VB.Form frmTabStripOptions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ceTabStrip - Options"
   ClientHeight    =   2625
   ClientLeft      =   10365
   ClientTop       =   3525
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkEnsureVisible 
      Caption         =   "Ensure associated tab is visible when a window is activated"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4695
   End
   Begin VB.CheckBox chkAlwaysMaximise 
      Caption         =   "Clicking a tab maximises the associated window"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   855
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.CheckBox chkFullWidth 
      Caption         =   "TabStrip width follows MDI width"
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.CheckBox chkMaintainVBW 
      Caption         =   "Maintain .vbw files"
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   6510
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3930
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   6510
   End
   Begin VB.CheckBox chkLogging 
      Caption         =   "Enable logging"
      Height          =   210
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Colin Edwards, 2021 "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "frmTabStripOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Sub chkAlwaysMaximise_Click()
   SaveSetting App.Title, "Options", "AlwaysMaximise", chkAlwaysMaximise.Value
   gAlwaysMaximise = chkAlwaysMaximise.Value = 1
End Sub
Private Sub chkEnsureVisible_Click()
   SaveSetting App.Title, "Options", "EnsureVisible", chkEnsureVisible.Value
   gEnsureVisible = chkEnsureVisible.Value = 1
End Sub
Private Sub chkFullWidth_Click()
   SaveSetting App.Title, "Options", "FullWidth", chkFullWidth.Value
   gShowFullWidth = chkFullWidth.Value = 0
End Sub
Private Sub chkLogging_Click()
   SaveSetting App.Title, "Options", "EnableLogging", chkMaintainVBW.Value
   gLogging = chkLogging.Value = 1
   Form_Resize
End Sub
Private Sub chkMaintainVBW_Click()
   SaveSetting App.Title, "Options", "MaintainVBW", chkMaintainVBW.Value
   gMaintainVBW = chkMaintainVBW.Value = 1
End Sub
Private Sub Command2_Click()
   List1.Clear
End Sub
Private Sub cmdClose_Click()
   Me.Hide
End Sub
Private Sub Form_Load()
   
   chkFullWidth.Value = GetSetting(App.Title, "Options", "FullWidth", 0)
   chkFullWidth.Enabled = Not gIsSDI
   chkFullWidth.Caption = chkFullWidth.Caption & IIf(gIsSDI, " (MDI only)", vbNullString)
   gShowFullWidth = chkFullWidth.Value = 0
   
   chkMaintainVBW.Value = GetSetting(App.Title, "Options", "MaintainVBW", 0)
   chkMaintainVBW.Enabled = Not gIsSDI
   chkMaintainVBW.Caption = chkMaintainVBW.Caption & IIf(gIsSDI, " (MDI only)", vbNullString)
   gMaintainVBW = chkMaintainVBW.Value = 1
   
   chkAlwaysMaximise.Value = GetSetting(App.Title, "Options", "AlwaysMaximise", 0)
   chkAlwaysMaximise.Enabled = Not gIsSDI
   chkAlwaysMaximise.Caption = chkAlwaysMaximise.Caption & IIf(gIsSDI, " (MDI only)", vbNullString)
   gAlwaysMaximise = chkAlwaysMaximise.Value = 1
   
   chkEnsureVisible.Value = GetSetting(App.Title, "Options", "EnsureVisible", 0)
   gEnsureVisible = chkEnsureVisible.Value = 1
   
   chkLogging.Value = GetSetting(App.Title, "Options", "EnableLogging", 0)
   gLogging = Me.chkLogging.Value = 1
   
   SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = UnloadMode = vbFormControlMenu
   If Cancel Then Me.Hide
End Sub
Public Sub Log(pText As String)
   List1.AddItem pText
   If List1.ListCount > 0 Then List1.ListIndex = List1.ListCount - 1
End Sub
Private Sub Form_Resize()
   With Me
      If chkLogging.Value = 1 Then
         .Move .Left, .Top, 6870, 7680
      Else
         .Move .Left, .Top, 4980, 3090
      End If
   End With
End Sub

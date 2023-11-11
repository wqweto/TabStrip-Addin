VERSION 5.00
Begin VB.Form frmTabStrip 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ceTabStrip"
   ClientHeight    =   1950
   ClientLeft      =   9990
   ClientTop       =   4575
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tWindowWatcher 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6480
      Top             =   1020
   End
   Begin TabStripAddIn.ucTabStrip ucTabStrip 
      Height          =   390
      Left            =   360
      Top             =   180
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   688
   End
   Begin VB.Label Label1 
      Caption         =   $"frmTabStrip.frx":0000
      Height          =   615
      Left            =   420
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuCloseDesigners 
         Caption         =   "Close all Designers"
      End
      Begin VB.Menu mnuCloseAllButThis 
         Caption         =   "All Except"
      End
      Begin VB.Menu mnuCloseToLeft 
         Caption         =   "All to the Left"
      End
      Begin VB.Menu mnuCloseToRight 
         Caption         =   "All to the Right"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGroupByProject 
         Caption         =   "Group by Project"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGroupByComponentType 
         Caption         =   "Group by Component Type"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestoreWorkspace 
         Caption         =   "Re-order per VBW file"
      End
   End
End
Attribute VB_Name = "frmTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMem4 Lib "msvbvm60" (src As Any, Dst As Any) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Const SW_NORMAL = 1
Const SW_MAXIMIZE = 1
Const WM_MDIMAXIMIZE As Long = &H225

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private WithEvents mWindowWatcher As cWindowWatcher
Attribute mWindowWatcher.VB_VarHelpID = -1
Private WithEvents mComponentEvents As VBIDE.VBComponentsEvents
Attribute mComponentEvents.VB_VarHelpID = -1
Private WithEvents mProjectEvents As VBIDE.VBProjectsEvents
Attribute mProjectEvents.VB_VarHelpID = -1
Private WithEvents mFileEvents As VBIDE.FileControlEvents
Attribute mFileEvents.VB_VarHelpID = -1
Private WithEvents mBuildEvents As VBIDE.VBBuildEvents
Attribute mBuildEvents.VB_VarHelpID = -1

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_SETFOCUS As Long = &H7
Const WM_CLOSE As Long = &H10
Const WM_SYSCOMMAND As Long = &H112
Const SC_CLOSE As Long = &HF060

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private mOldParentHwnd As Long
Public Sub Init()
Dim objEvents2 As Events2
   Set mComponentEvents = gVBInstance.Events.VBComponentsEvents(Nothing)
   Set mProjectEvents = gVBInstance.Events.VBProjectsEvents
   Set mFileEvents = gVBInstance.Events.FileControlEvents(Nothing)
   Set objEvents2 = gVBInstance.Events
   Set mBuildEvents = objEvents2.VBBuildEvents
   mOldParentHwnd = SetParent(ucTabStrip.hWnd, gToolbarHwnd)
   Set mWindowWatcher = New cWindowWatcher
   ucTabStrip.LockUpdate = True
      mWindowWatcher.Init
      If gLaunchedAtIDEStartup And Not gVBInstance.ActiveVBProject Is Nothing Then mnuRestoreWorkspace_Click
   ucTabStrip.LockUpdate = False
   tWindowWatcher.Interval = 125
   tWindowWatcher.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set mWindowWatcher = Nothing
   Set mComponentEvents = Nothing
   Set mProjectEvents = Nothing
   Set mFileEvents = Nothing
   Set mBuildEvents = Nothing
   SetParent ucTabStrip.hWnd, mOldParentHwnd
End Sub
Private Sub mComponentEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
   If gLogging Then Log "Component '" & OldName & "' was renamed to '" & VBComponent.Name
   RenameTabs VBComponent.Collection.Parent, VBComponent
End Sub

'========================================
'SIGNIFICANT EVENTS
' All events here suggest potentially relevant IDE activity, and so represent a good time to forcibly update our TabStrip, rather than waiting on the polling timer
Private Sub mBuildEvents_EnterRunMode()
   If gLogging Then Log "Entering Run Mode"
   ucTabStrip.LockUpdate = True
   RefreshWindows True
   ucTabStrip.LockUpdate = False
End Sub
Private Sub mBuildEvents_EnterDesignMode()
   If gLogging Then Log "Entering Design Mode"
   RefreshWindows True
End Sub
Private Sub mProjectEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
   If gLogging Then Log "Project '" & VBProject.Name & "' has completed loading: IDE has " & gVBInstance.Windows.Count & " windows" ' also see mFileEvents_BeforeLoadFile
   RefreshWindows True
End Sub
Private Sub mComponentEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
   If gLogging Then Log "Component '" & VBComponent.Name & "' was added"
   RefreshWindows True
End Sub
Private Sub mComponentEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
   If gLogging Then Log "Component '" & VBComponent.Name & "' was removed"
   RefreshWindows True
End Sub
'END SIGNIFICANT EVENTS
'----------------------------

Private Sub RenameTabs(ByVal Project As VBProject, ByVal Component As VBComponent)
Dim CP As CodePane, VBComp As VBComponent, VBProj As VBProject, hWnd As Long

   On Error GoTo ErrHandler
   
   For Each CP In gVBInstance.CodePanes
      Set VBComp = CP.CodeModule.Parent
      Set VBProj = VBComp.Collection.Parent
      If VBProj.Name = Project.Name And VBComp.Name = Component.Name Then
         GetMem4 ByVal ObjPtr(CP.Window) + &H1C, hWnd
         ucTabStrip.SetCaption "W" & hWnd, Component.Name
      End If
   Next CP
   
   If Component.HasOpenDesigner Then
      GetMem4 ByVal ObjPtr(Component.DesignerWindow) + &H1C, hWnd
      ucTabStrip.SetCaption "W" & hWnd, Component.Name & " (Design)"
   End If
   
   Exit Sub
   
ErrHandler:
   If gLogging Then Log "RenameTabs: " & Err.Description
End Sub
'WindowWatcher Events
Private Sub mWindowWatcher_BatchChangeOccurring()
'We're about to get multiple messages - don't want to slow things down with animations and the like
   If gLogging Then Log "Bulk change occurring - suspending TabStrip"
   ucTabStrip.LockUpdate = True
End Sub
Private Sub mWindowWatcher_BulkChangeComplete()
   If gLogging Then Log "Bulk change complete - unsuspending TabStrip"
   ucTabStrip.LockUpdate = False
End Sub
Private Sub mWindowWatcher_MDIWindowSizeChanged(ByVal NewX As Long, ByVal NewW As Long)
   ucTabStrip.Move NewX, 1, NewW, TOOLBAR_HEIGHT
   If gLogging Then Log "Move Toolbar to x=" & NewX & ", w=" & NewW
End Sub
Private Sub mWindowWatcher_ActiveWindowChanged(ByVal NewActivehWnd As Long)
   If ucTabStrip.ActiveTabKey = "W" & NewActivehWnd Then Exit Sub
   If gLogging Then Log "'" & CaptionForHwnd(NewActivehWnd) & "' Activated via WindowWatcher"
   ucTabStrip.SetActiveTab "W" & NewActivehWnd, gEnsureVisible
End Sub
Private Sub mWindowWatcher_WindowAdded(ByVal hWnd As Long, ByVal TabName As String, ByVal IconHandle As Long)
   If gLogging Then Log "'" & CaptionForHwnd(hWnd) & "' Activated as a new window"
   ucTabStrip.AddTab TabName, "W" & hWnd, IconHandle
   'ucTabStrip.SetActiveTab "W" & hWnd
End Sub
Private Sub mWindowWatcher_WindowRemoved(ByVal hWnd As Long)
   If gLogging Then Log "'" & hWnd & "' no longer exists per the WindowWatcher"
   If ucTabStrip.TabExists("W" & hWnd) Then ucTabStrip.RemoveTab "W" & hWnd, "W" & mWindowWatcher.GetActiveWindowhWnd
End Sub
'UI Elements
Private Sub ucTabStrip_TabClick(ByVal TabKey As String, ByVal WasAlreadyActive As Boolean)
Dim hWnd As Long
   hWnd = CLng(Replace(TabKey, "W", vbNullString))
   If gLogging Then Log "'" & CaptionForHwnd(hWnd) & " clicked by user" & IIf(WasAlreadyActive, " (Tab was already active)", ": Sending activate msg")
   If gIsSDI Then
      If IsIconic(hWnd) Then
         ShowWindow hWnd, SW_NORMAL
      Else
         If Not WasAlreadyActive Then SetActiveWindow hWnd
      End If
   Else
      If gAlwaysMaximise Then PostMessage gMDIhWnd, WM_MDIMAXIMIZE, hWnd&, 0&
      If Not WasAlreadyActive Then PostMessage hWnd, WM_SETFOCUS, 0&, 0&
   End If
End Sub
Private Sub ucTabStrip_TabRightClick(ByVal TabKey As String)
Dim sTabCaption As String
   sTabCaption = ucTabStrip.Caption(TabKey)
   mnuCloseAllButThis.Caption = "Close all except '" & sTabCaption & "'"
   mnuCloseToLeft.Caption = "Close all left of '" & sTabCaption & "'"
   mnuCloseToRight.Caption = "Close all right of '" & sTabCaption & "'"
   
   mnuActions.Tag = TabKey 'this will be the Tab that the bulk close centres around
   ucTabStrip.LockUpdate = True 'keeps mouse-over effect on the tab that was right-clicked on
      PopupMenu mnuActions
   ucTabStrip.LockUpdate = False
End Sub
Private Sub ucTabStrip_TabCloseClick(ByVal TabKey As String, ByRef Cancel As Boolean)
Dim hWnd As Long
   'here, we'll cancel the user click and let the WindowWatcher do the rest in response to WM_CLOSE
   hWnd = Replace(TabKey, "W", vbNullString)
   If gLogging Then Log "User-click on Tab Close button - sending WM_CLOSE to '" & CaptionForHwnd(hWnd) & "'"
   
   If gIsSDI And ClassName(hWnd) <> "VbaWindow" And ClassName(hWnd) <> "DesignerWindow" Then
      SendMessage hWnd, WM_SYSCOMMAND, SC_CLOSE, 0&
   Else
      SendMessage hWnd, WM_CLOSE, 0&, 0&
   End If
   Cancel = True
   RefreshWindows True 'force our cache of windows to update
End Sub
Private Sub tWindowWatcher_Timer()
   RefreshWindows
End Sub
Private Sub RefreshWindows(Optional Force As Boolean)
   tWindowWatcher.Enabled = False
      mWindowWatcher.Refresh Force
   tWindowWatcher.Enabled = True
End Sub

'========================================
'WORKSPACE MAINTENANCE - see ModWorkspaceVBW
Private Sub mFileEvents_RequestWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileName As String, Cancel As Boolean)
   If Not Right$(FileName, 4) = ".vbp" Then Exit Sub
   If gIsSDI Then Exit Sub
   If Len(VBProject.FileName) = 0 Then Exit Sub
   If gMaintainVBW Then SaveVBWFile VBProject
End Sub
Private Sub mnuRestoreWorkspace_Click()
Dim VBProj As VBProject
   For Each VBProj In gVBInstance.VBProjects
      RestoreFromVBW VBProj
   Next VBProj
End Sub
'END WORKSPACE MAINTENANCE
'----------------------------

'=====================================
'BULK TAB CLOSURE
' All the following code relates to the menu options for the above
Private Sub mnuCloseAll_Click()
   BulkCloseTabs "All"
End Sub
Private Sub mnuCloseAllButThis_Click()
   BulkCloseTabs "AllButThis"
End Sub
Private Sub mnuCloseDesigners_Click()
   BulkCloseTabs "Designers"
End Sub
Private Sub mnuCloseToLeft_Click()
   BulkCloseTabs "Left"
End Sub
Private Sub mnuCloseToRight_Click()
   BulkCloseTabs "Right"
End Sub
Private Sub BulkCloseTabs(CloseOption As String) 'supporting sub for the menu 'bulk close' options
Dim CP As CodePane, VBComp As VBComponent, VBProj As VBProject, hWnd As Long, sVBWInfo As String, TargetWindows As Dictionary, Pos As Long, i As Long

   tWindowWatcher.Enabled = False
   ucTabStrip.LockUpdate = True
   
   Set TargetWindows = New Dictionary
   
   Select Case CloseOption
      Case "Right"
         Pos = ucTabStrip.TabPosition(mnuActions.Tag)
         For i = Pos + 1 To ucTabStrip.TabCount
            TargetWindows.Add ucTabStrip.KeyFromPosition(i), 0
         Next i
      Case "Left"
         If gLogging Then Log "Closing to the left of " & ucTabStrip.Caption(mnuActions.Tag) & " in position " & ucTabStrip.TabPosition(mnuActions.Tag)
         Pos = ucTabStrip.TabPosition(mnuActions.Tag)
         For i = 1 To Pos - 1
            TargetWindows.Add ucTabStrip.KeyFromPosition(i), 0
         Next i
      Case "AllButThis"
         TargetWindows.Add mnuActions.Tag, 0
   End Select
   
   If Not CloseOption = "Designers" Then
      For Each CP In gVBInstance.CodePanes
         If CloseOption = "All" Then
            CP.Window.Close
         Else
            GetMem4 ByVal ObjPtr(CP.Window) + &H1C, hWnd
            If CloseOption = "AllButThis" Then
               If Not TargetWindows.Exists("W" & hWnd) Then CP.Window.Close
            Else
               If TargetWindows.Exists("W" & hWnd) Then CP.Window.Close
            End If
         End If
      Next CP
   End If
   
   For Each VBProj In gVBInstance.VBProjects
      For Each VBComp In VBProj.VBComponents
         If VBComp.HasOpenDesigner Then
            If CloseOption = "Designers" Or CloseOption = "All" Then
               VBComp.DesignerWindow.Close
            Else
               GetMem4 ByVal ObjPtr(VBComp.DesignerWindow) + &H1C, hWnd
               If CloseOption = "AllButThis" Then
                  If Not TargetWindows.Exists("W" & hWnd) Then VBComp.DesignerWindow.Close
               Else
                  If TargetWindows.Exists("W" & hWnd) Then VBComp.DesignerWindow.Close
               End If
            End If
         End If
      Next VBComp
   Next VBProj
   
   RefreshWindows True
   ucTabStrip.LockUpdate = False
   
End Sub
'END BULK TAB CLOSURE
'----------------------


'=====================================
'INFORMATION EVENTS
'These are just here for studying in the log, at the moment! We don't use any of them...
Private Sub mFileEvents_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
   If gLogging Then Log "File removed from '" & VBProject.Name & "': " & FileName & "(" & FileType & ")"
End Sub
Private Sub mFileEvents_BeforeLoadFile(ByVal VBProject As VBIDE.VBProject, FileNames() As String)
   If gLogging And Right$(FileNames(0), 4) = ".vbp" Then Log "Begin loading " & FileNames(0)
End Sub
Private Sub mFileEvents_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
   If gLogging Then Log "File added to '" & VBProject.Name & "': " & FileName & "(" & FileType & ")"
End Sub
'These are paired here because these two events mark the beginning and end of a project's removal - may come in handy...
Private Sub mProjectEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
   If gLogging Then Log "Project '" & VBProject.Name & "' is being removed"
End Sub
Private Sub mFileEvents_AfterCloseFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal WasDirty As Boolean)
   If gLogging And FileType = vbext_ft_Project Then Log "Project '" & VBProject.Name & "' was closed"
End Sub
'END INFORMATION EVENTS
'----------------------


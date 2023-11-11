VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindowWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_MDIGETACTIVE As Long = &H229
Private Const ICON_SMALL As Long = 0
Private Const ICON_BIG As Long = 1
Private Const ICON_SMALL2 As Long = 2
Private Const WM_GETICON As Long = &H7F

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GCL_HICONSM = (-34)
Private Const GCL_HICON = (-14)

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private mCurrentWindows As Dictionary
Private mUCx As Long, mUCw As Long
Private mActiveWindowhWnd As Long
Private mVBWindowCount As Long
Private mActiveWindowChanged As Boolean

Event BatchChangeOccurring() 'if we spot that we've got lots of window changes occurring, we'll broadcast it
Event BulkChangeComplete()
Event ActiveWindowChanged(ByVal NewActivehWnd As Long)
Event WindowAdded(ByVal hWnd As Long, ByVal TabName As String, ByVal IconHandle As Long)
Event WindowRemoved(ByVal hWnd As Long)
Event MDIWindowSizeChanged(ByVal NewX As Long, ByVal NewW As Long)
Public Sub Init()
   Set mCurrentWindows = New Dictionary
   Refresh True
End Sub
Public Sub Refresh(Optional Forced As Boolean)
   
   If MDIWindowSizeChanged Then RaiseEvent MDIWindowSizeChanged(mUCx, mUCw)
   
   If Not QuickTestsPassed Or Forced Then
      'we only enumerate all the windows when it seems neccessary to do so
      CheckForWindowChanges
      If mActiveWindowChanged Then RaiseEvent ActiveWindowChanged(mActiveWindowhWnd)
   End If

End Sub
Public Function GetActiveWindowhWnd() As Long
      
   On Error GoTo ErrHandler
   
   If gIsSDI Then
      GetActiveWindowhWnd = GetActiveWindow
   Else
      GetActiveWindowhWnd = SendMessage(gMDIhWnd, WM_MDIGETACTIVE, 0&, ByVal 0&)
   End If

   Exit Function

ErrHandler:
   If gLogging Then Log "GetActiveWindowhWnd: " & Err.Description
End Function
Private Sub CheckForWindowChanges()
Dim thisKey As Variant, RefreshedWindows As Dictionary, sTabCaption As String, sWindowClass As String
Dim BulkChangeOccurring As Boolean

   On Error GoTo error_handler
   
   Set RefreshedWindows = EnumIDEWindows
   
   BulkChangeOccurring = Abs(RefreshedWindows.Count - mCurrentWindows.Count) > 3
   If BulkChangeOccurring Then RaiseEvent BatchChangeOccurring
   
   'first, remove any closed windows
   For Each thisKey In mCurrentWindows
      If Not RefreshedWindows.Exists(thisKey) Then
         mCurrentWindows.Remove thisKey
         RaiseEvent WindowRemoved(CLng(thisKey))
      End If
   Next thisKey
   
   'then, check for new ones
   For Each thisKey In RefreshedWindows
      If Not mCurrentWindows.Exists(thisKey) Then
         sWindowClass = RefreshedWindows(thisKey)
         sTabCaption = TabCaptionForhWnd(CLng(thisKey), sWindowClass)
         mCurrentWindows.Add thisKey, sWindowClass
         RaiseEvent WindowAdded(CLng(thisKey), sTabCaption, IconForWindow(CLng(thisKey)))
      End If
   Next thisKey
   
   If BulkChangeOccurring Then RaiseEvent BulkChangeComplete
   
   Exit Sub
   
error_handler:
   If gLogging Then Log "CheckForWindowChanges: " & Err.Description
   RaiseEvent BulkChangeComplete 'won't do any harm - just unsuspends the TabStrip
End Sub
Private Function QuickTestsPassed() As Boolean
Dim hWnd As Long, CurrentVBWindowCount As Long, WindowCountChanged As Boolean
   
   hWnd = GetActiveWindowhWnd
   mActiveWindowChanged = (mActiveWindowhWnd <> hWnd)
   mActiveWindowhWnd = hWnd
   
   CurrentVBWindowCount = gVBInstance.Windows.Count
   WindowCountChanged = CurrentVBWindowCount <> mVBWindowCount
   mVBWindowCount = CurrentVBWindowCount

   QuickTestsPassed = (Not mActiveWindowChanged) And (Not WindowCountChanged)

End Function
Private Function MDIWindowSizeChanged() As Boolean
Dim R As RECT, P As POINTAPI
   If gIsSDI Or gShowFullWidth Then
      GetClientRect gIDEhWnd, R
      MDIWindowSizeChanged = (mUCw <> R.Right)
      mUCx = 0: mUCw = R.Right
   Else
      GetWindowRect gMDIhWnd, R
      P.x = R.Left: P.y = R.Top
      ScreenToClient gIDEhWnd, P
      MDIWindowSizeChanged = (mUCx < P.x) Or (mUCw <> R.Right - R.Left)
      mUCx = P.x: mUCw = R.Right - R.Left
   End If
End Function
Private Function IconForWindow(ByVal hWnd As Long) As Long 'only called upon discovery (or rediscovery) of a new window
   IconForWindow = SendMessage(hWnd, WM_GETICON, ICON_SMALL, 0)
   If IconForWindow = 0 Then IconForWindow = GetClassLong(hWnd, GCL_HICONSM)
   If IconForWindow = 0 Then IconForWindow = SendMessage(hWnd, WM_GETICON, ICON_SMALL2, 0)
   If IconForWindow = 0 Then IconForWindow = SendMessage(hWnd, WM_GETICON, ICON_BIG, 0)
   If IconForWindow = 0 Then IconForWindow = GetClassLong(hWnd, GCL_HICON)
End Function
Private Function TabCaptionForhWnd(ByVal hWnd As Long, ByVal ClassName As String) 'only called upon discovery (or rediscovery) of a new window
Dim s() As String, RemoveProjectPrefix As Boolean, AddDesignSuffix As Boolean
   
   TabCaptionForhWnd = CaptionForHwnd(hWnd)
   
   Select Case ClassName
      Case "DockingView"
      Case "ThunderForm", "ThunderMDIForm" 'only apply to SDI
         AddDesignSuffix = True
      Case "VbaWindow", "DesignerWindow", "ThunderDFrame" 'the latter only applies to SDI
         AddDesignSuffix = ClassName <> "VbaWindow": RemoveProjectPrefix = True
         s = Split(TabCaptionForhWnd, " (") 'strip whatever suffix is currently present
         TabCaptionForhWnd = s(0)
   End Select
   
   If RemoveProjectPrefix Then
      s = Split(TabCaptionForhWnd, " - ")
      TabCaptionForhWnd = s(UBound(s))
   End If
   
   If AddDesignSuffix Then TabCaptionForhWnd = TabCaptionForhWnd & " (Design)"
   
End Function
Private Sub Class_Terminate()
   Set mCurrentWindows = Nothing
End Sub

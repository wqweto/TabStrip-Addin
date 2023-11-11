VERSION 5.00
Begin VB.UserControl ucTabStrip 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   Begin VB.Timer tMouseTimer 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3840
      Top             =   120
   End
   Begin VB.Label lblScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   7680
      TabIndex        =   1
      Top             =   120
      Width           =   270
   End
   Begin VB.Label lblScroll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
End
Attribute VB_Name = "ucTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
   x  As Long
   y  As Long
End Type

Private Declare Function GetCapture Lib "user32" () As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y2 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal W As Long, ByVal H As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal iStepIfAniCur As Long, ByVal hBrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL As Long = &H3

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Const DT_LEFT As Long = &H0&
Const DT_VCENTER As Long = &H4&
Const DT_SINGLELINE As Long = &H20&
Const DT_END_ELLIPSIS As Long = &H8000&

Private Type TabItem
   Key As String 'unqiue ID for the tab
   Caption As String
   HomeX As Long 'a tabs 'home' x co-ord; a new 'home' is assigned if the tabs are re-ordered (e.g. by the user dragging to re-arrange them)
   x As Long 'a tabs current x-position, equal to HomeX, unless it's currently being influenced by a user re-arrangement (i.e. drag)
   Width As Long 'what it says
   Hidden As Boolean 'tab exists, but is not visible to the user
   IconHandle As Long
End Type

Private Tabs() As TabItem

Private Const ICON_SPACE_W As Long = 18
Private Const TAB_H As Long = 26
Private Const BTTN_W As Long = 20

Private mPositions As Collection 'the current positions of the tabs: e.g. if mPositions(3) = 5, then Tabs(5) is in position 3 (from L to R)
Private mViewportX As Long
Private mMaxViewportX As Long
Private mViewportW As Long
Private mTotalTabsW As Long
Private mMouseDownX As Long
Private mScrollDragging As Boolean 'true when the user is drag-scrolling (via CTRL-down) the whole UC
Private mReOrderDragging As Boolean 'true when the user is dragging to re-arrange the tabs
Private mAutoScrollDirection As Long '-1,0,or +1, depending on whether the timer is auto-scrolling the control (sign = direction)
Private mDragDistance As Long 'only used to work out the drag direction
Private mDragDirection As Long 'drag direction
Private mDragFromPosition As Long
Private mDragToPosition As Long
Private mDraggedTabIndex As Long
Private mMouseOverTabIndex As Long
Private mActiveTabIndex As Long
Private mCloseButtonIsHot As Boolean
Private mMouseUpButton As MouseButtonConstants
Private mLockUpdate  As Boolean
Private mHotButtonIndex As Long 'the index of the scroll button the mouse is over (becomes negative to indicate mouse down)
Private mButtonWidth As Long 'can be BTTN_W or 0 (0 when the buttons aren't needed)

Event TabRightClick(ByVal TabKey As String)
Event TabClick(ByVal TabKey As String, ByVal WasAlreadyActive As Boolean)
Event TabCloseClick(ByVal TabKey As String, ByRef Cancel As Boolean)
Public Property Let LockUpdate(TrueFalse As Boolean)
   mLockUpdate = TrueFalse
   If Not mLockUpdate Then Init
End Property
Public Property Get Caption(ByVal Key As String) As String
   Caption = Tabs(mPositions(Key)).Caption
End Property
Public Function KeyFromPosition(pPosition As Long) As String
   KeyFromPosition = Tabs(mPositions(pPosition)).Key
End Function
Public Property Get TabPosition(Key As String) As String
   TabPosition = PositionForTabIndex(mPositions(Key))
End Property
Public Property Get TabCount() As Long
   TabCount = mPositions.Count
End Property
Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property
Public Property Get ActiveTabKey() As String
   If mActiveTabIndex = 0 Then ActiveTabKey = vbNullString Else ActiveTabKey = Tabs(mActiveTabIndex).Key
End Property
Public Function TabExists(pKey As String) As Boolean
   On Error Resume Next
   IsObject mPositions.Item(pKey)
   TabExists = Err.Number = 0
End Function
Public Sub SetActiveTab(Key As String, Optional EnsureVisible As Boolean)
   
   If Not TabExists(Key) Then Exit Sub
   
   mActiveTabIndex = mPositions(Key)
   
   If mLockUpdate Then Exit Sub
   
   If EnsureVisible Then EnsureTabVisible Key
   
   Refresh

End Sub
Public Sub AddTab(ByVal pCaption As String, ByVal pKey As String, Optional pIconHandle As Long, Optional pPosition As Long)
   ReDim Preserve Tabs(UBound(Tabs) + 1)
   If pPosition = 0 Then pPosition = UBound(Tabs)
   Tabs(pPosition).Caption = pCaption
   Tabs(pPosition).Key = pKey
   UserControl.FontBold = False
   Tabs(pPosition).Width = UserControl.TextWidth(pCaption) + ICON_SPACE_W + 12
   Tabs(pPosition).IconHandle = pIconHandle
   mPositions.Add pPosition, pKey
   mTotalTabsW = mTotalTabsW + Tabs(pPosition).Width
   ReCalcMetrics
   If Not mLockUpdate Then Init pKey Else SetActiveTab pKey
End Sub
Public Sub SetCaption(pKey As String, pCaption As String)
Dim Idx As Long
   If Not TabExists(pKey) Then Exit Sub
   
   Idx = mPositions(pKey)
   Tabs(Idx).Caption = pCaption
   Tabs(Idx).Width = UserControl.TextWidth(pCaption) + ICON_SPACE_W + 12
   Init
End Sub
Public Sub RemoveTab(ByVal Key As String, Optional NewActiveTabKey As String) 'when called from 'outside', this is always silent (i.e. no event raised)
Dim RemovedIdx As Long, thisPos As Long, thisTabKey As String, thisTabIndx As Long, ActiveTabKey As String
   
   If Not TabExists(Key) Then Exit Sub
   
   RemovedIdx = mPositions(Key)
   thisPos = PositionForTabIndex(RemovedIdx)
   If NewActiveTabKey = Key Then NewActiveTabKey = vbNullString 'can't do this - makes no sense!
   If Not TabExists(NewActiveTabKey) Then NewActiveTabKey = vbNullString 'can't do this - makes no sense!
   
   'we activate a new tab BEFORE removal, so that the RemoveTab animation looks better! ;)
   If Len(NewActiveTabKey) Then  'we've been told what tab to make active when this one closes
      SetActiveTab NewActiveTabKey
   Else 'if not we may need to work it out...
      If RemovedIdx = mActiveTabIndex Then '...but only if we're about to close the active tab
         If mPositions.Count = 1 Then
            mActiveTabIndex = 0 'we're about to close the last remaining tab
         ElseIf thisPos > 1 Then
            mActiveTabIndex = mPositions(thisPos - 1)
         Else
            mActiveTabIndex = mPositions(thisPos + 1)
         End If
         SetActiveTab Tabs(mActiveTabIndex).Key
      End If
   End If
   
   ActiveTabKey = Tabs(mActiveTabIndex).Key 'store the active key as it's about to get a new index!
   
   DoRemoveEffect thisPos, RemovedIdx 'but first, a little animation effect
   ReCalcMetrics
   
   mPositions.Remove Key
   
   'Loop thru mPositions, removing and re-adding any item where the Tab index is <= the one we're removing. Then re-add, using one number lower for the Tab index
   For thisPos = 1 To mPositions.Count
      thisTabIndx = mPositions(thisPos)
      If thisTabIndx > RemovedIdx Then
         thisTabKey = Tabs(thisTabIndx).Key
         mPositions.Remove thisTabKey
         If mPositions.Count = 0 Then 'we just removed the last item 1 line up!
            mPositions.Add 1, thisTabKey
         ElseIf thisPos = 1 Then
            mPositions.Add thisTabIndx - 1, thisTabKey, 1
         Else
            mPositions.Add thisTabIndx - 1, thisTabKey, , thisPos - 1
         End If
      End If
   Next thisPos
   
   For thisTabIndx = RemovedIdx To UBound(Tabs) - 1
      Tabs(thisTabIndx) = Tabs(thisTabIndx + 1)
   Next thisTabIndx
   ReDim Preserve Tabs(UBound(Tabs) - 1)
   
   If mActiveTabIndex > 0 Then mActiveTabIndex = mPositions(ActiveTabKey) 'ascertain where our active tab is now
   
   If Not mLockUpdate Then Init

End Sub
Public Sub Init(Optional ByVal ActiveTabKey As String)
Dim i As Long, thisIndex As Long, prevIndex As Long
   
   If UBound(Tabs) > 0 Then
      Tabs(mPositions(1)).x = 0
      For i = 1 To UBound(Tabs)
         thisIndex = mPositions(i)
         If i > 1 Then
            prevIndex = mPositions(i - 1)
            Tabs(thisIndex).x = Tabs(prevIndex).x + Tabs(prevIndex).Width '- 1
         End If
         Tabs(thisIndex).HomeX = Tabs(thisIndex).x
      Next i
      mTotalTabsW = Tabs(mPositions(mPositions.Count)).x + Tabs(mPositions(mPositions.Count)).Width
   Else
      ActiveTabKey = vbNullString 'ignore it if it was passed - we have no tabs!
   End If
   
   ReCalcMetrics
   If Len(ActiveTabKey) Then SetActiveTab ActiveTabKey, True Else Refresh

End Sub
Private Sub DoDragReOrdering(Direction As Long)
Dim Pos As Long, Idx As Long

   If Direction = 0 Then Exit Sub
   
   If Direction = 1 Then mDragToPosition = 1 Else mDragToPosition = mPositions.Count
   
   For Pos = 1 To UBound(Tabs)
      Idx = mPositions(Pos)
      If Idx <> mDraggedTabIndex Then
         If Direction = 1 Then 'right
            If (Tabs(mDraggedTabIndex).x + Tabs(mDraggedTabIndex).Width) >= Tabs(Idx).HomeX + Tabs(Idx).Width / 2 Then
               mDragToPosition = mDragToPosition + 1
               If mDragFromPosition < Pos Then Tabs(Idx).HomeX = Tabs(Idx).HomeX - Tabs(mDraggedTabIndex).Width
            End If
         ElseIf Direction = -1 Then 'left
            If (Tabs(mDraggedTabIndex).x) <= Tabs(Idx).HomeX + Tabs(Idx).Width / 2 Then
               mDragToPosition = mDragToPosition - 1
               If mDragFromPosition > Pos Then Tabs(Idx).HomeX = Tabs(Idx).HomeX + Tabs(mDraggedTabIndex).Width
            End If
         End If
      End If
   Next Pos
   
   If mDragToPosition <> mDragFromPosition Then
      MoveTab Tabs(mDraggedTabIndex).Key, mDragFromPosition, mDragToPosition
      mDragFromPosition = mDragToPosition
   End If
   
   Refresh

End Sub
Private Function ParkTabs(ExcludeIndex As Long, Optional EffectIncrements As Long = 16) As Boolean
Dim Idx As Long, ParkedTabCount As Long
   
   For Idx = 1 To UBound(Tabs)
      If Idx <> ExcludeIndex Then 'ExcludeIndex is usually the one being dragged
         If Abs(Tabs(Idx).x - Tabs(Idx).HomeX) > EffectIncrements Then
            Tabs(Idx).x = Tabs(Idx).x + Sgn(Tabs(Idx).HomeX - Tabs(Idx).x) * EffectIncrements
         Else
            Tabs(Idx).x = Tabs(Idx).HomeX
            ParkedTabCount = ParkedTabCount + 1
         End If
      End If
   Next Idx
   
   ParkTabs = (ParkedTabCount = mPositions.Count)

   Refresh

End Function
Private Sub DoRemoveEffect(RemovedPos As Long, Idx As Long)
Dim i As Long, Increment As Long, Remainder As Long
   
   If RemovedPos = mPositions.Count And mViewportX = 0 Then Exit Sub

   Tabs(Idx).Hidden = True
   For i = RemovedPos To mPositions.Count
      Tabs(mPositions(i)).HomeX = Tabs(mPositions(i)).HomeX - Tabs(Idx).Width
   Next i
   
   If Not mLockUpdate Then
      If Tabs(Idx).x > UserControl.ScaleWidth + mViewportX Then
         mTotalTabsW = mTotalTabsW - Tabs(Idx).Width 'the removed tab isn't visible - early bail
      ElseIf Tabs(Idx).x + Tabs(Idx).Width < mViewportX Then
         mTotalTabsW = mTotalTabsW - Tabs(Idx).Width 'the removed tab isn't visible - early bail
      Else
         Increment = Tabs(Idx).Width \ 8
         Remainder = Tabs(Idx).Width Mod 8
         For i = 1 To 9
            If i = 9 Then Increment = Remainder
            mTotalTabsW = mTotalTabsW - Increment
            If mTotalTabsW - mViewportX <= UserControl.ScaleWidth Then ShiftViewport -Increment, True
            ParkTabs 0, Increment 'let's exploit the animation effect we created for the tab dragging...
            Sleep 10
         Next i
      End If
   Else
      mTotalTabsW = mTotalTabsW - Tabs(Idx).Width
      If mViewportX > 0 Then ShiftViewport -Tabs(Idx).Width
   End If
End Sub
Public Sub MoveTab(ByVal Key As String, Optional ByVal FromPosition As Long, Optional ByVal ToPosition As Long)
Dim TabIdx As Long
   
   TabIdx = mPositions(Key)
   
   If FromPosition = 0 Then FromPosition = PositionForTabIndex(TabIdx)
   If ToPosition = 0 Then ToPosition = mPositions.Count
   If FromPosition = ToPosition Then Exit Sub
   
   mPositions.Remove Key
   
   If FromPosition < ToPosition Then
      mPositions.Add TabIdx, Key, , ToPosition - 1
   Else
      mPositions.Add TabIdx, Key, ToPosition
   End If
End Sub
Public Sub EnsureTabVisible(ByVal Key As String)
Dim DX As Double, prevDX As Double, i As Double, TotalDX As Double, prevTotalDX As Double
Dim thisTab As TabItem, ScrollDistance As Long, Steps As Long
   
   thisTab = Tabs(mPositions(Key))
   If thisTab.x < mViewportX Then
      ScrollDistance = thisTab.x - mViewportX
   ElseIf thisTab.x + thisTab.Width > mViewportX + mViewportW Then
      ScrollDistance = (thisTab.x + thisTab.Width) - (mViewportX + mViewportW)
   Else
      ShiftViewport 0 'in case we had a control resize
      Exit Sub
   End If
   
   If mLockUpdate Then
      ShiftViewport ScrollDistance, True
      Exit Sub
   End If
   
   Steps = 2 + Abs(ScrollDistance) / 30
   If Steps > 30 Then Steps = 24
   
   DX = 1
   For i = -1 + 1 / Steps To 1 Step 1 / Steps
      prevDX = DX
      prevTotalDX = TotalDX
      DX = Sqr(1 - (1 - Abs(i)) ^ 2)
      TotalDX = TotalDX + Abs(prevDX - DX) * ScrollDistance / 2
      ShiftViewport CLng(TotalDX) - CLng(prevTotalDX)
      Sleep 10
   Next i

End Sub
Private Sub lblScroll_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      mHotButtonIndex = -Index 'when the sign is flipped, we scroll (the timer picks this up and acts on it)
   Else
      If Index = 1 Then EnsureTabVisible Tabs(mPositions(1)).Key Else EnsureTabVisible Tabs(mPositions(mPositions.Count)).Key
   End If
End Sub
Private Sub lblScroll_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim R As RECT
   If mMouseOverTabIndex > 0 Then mMouseOverTabIndex = 0: Refresh
   mCloseButtonIsHot = False
   
   If Not tMouseTimer.Enabled Then tMouseTimer_Timer 'turn on mouse-leave checking
   If Index = 1 And mViewportX = 0 Then Exit Sub
   If Index = 2 And mViewportX = mMaxViewportX Then Exit Sub
   
   If Button = vbLeftButton Then
      'allows the user to move the mouse on and off the button to pause the scroll
      R.Right = lblScroll(Index).Width: R.Bottom = lblScroll(Index).Height
      If PtInRect(R, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY) Then
         Index = -Index
      Else
         Index = 0
      End If
   End If
   
   If mHotButtonIndex <> Index Then
      mHotButtonIndex = Index
      Refresh
   End If
End Sub
Private Sub lblScroll_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   mHotButtonIndex = Abs(mHotButtonIndex)
   Refresh
End Sub
Private Sub UserControl_Click()
Dim Key As String, Cancel As Boolean, prevIndex As Long
   
   If mMouseOverTabIndex = 0 Then Exit Sub
   
   Key = Tabs(mMouseOverTabIndex).Key
   
   If mCloseButtonIsHot Then 'if this is true then we must have a mMouseOverTabIndex>0
      mCloseButtonIsHot = False
      mMouseOverTabIndex = 0
      RaiseEvent TabCloseClick(Key, Cancel)
      If Not Cancel Then RemoveTab Key
      'Init NextActiveTabKey
   Else
      If mMouseUpButton = vbLeftButton Then
         prevIndex = mActiveTabIndex
         If Key <> Tabs(mActiveTabIndex).Key Then SetActiveTab Key, True Else EnsureTabVisible Key
         RaiseEvent TabClick(Key, prevIndex = mActiveTabIndex)
      ElseIf mMouseUpButton = vbRightButton Then
         RaiseEvent TabRightClick(Key)
      End If
   End If

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton And mMouseOverTabIndex > 0 Then mMouseDownX = mViewportX + x
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim DragDistance As Long, TabIdx As Long, OverCloseButton As Boolean, DoRefresh As Boolean, BeginDrag As Boolean
   
   mHotButtonIndex = 0
   x = x + mViewportX
   If Button = 0 Then
      If Not tMouseTimer.Enabled Then tMouseTimer_Timer 'turn on mouse-leave checking
      TabIdx = HitTest(x - mButtonWidth)
      If TabIdx <> mMouseOverTabIndex Then
         mMouseOverTabIndex = TabIdx
         DoRefresh = True
      End If
      OverCloseButton = CloseButtonHitTest(x - mButtonWidth, y)
      If mCloseButtonIsHot <> OverCloseButton Then
         mCloseButtonIsHot = OverCloseButton
         DoRefresh = True
      End If
      If DoRefresh Then Refresh
   ElseIf Button = vbLeftButton And mMouseDownX <> -1 Then
      
      DragDistance = x - mMouseDownX
      
      If Not mReOrderDragging And Not mScrollDragging Then 'begin drag?
         BeginDrag = (Abs(DragDistance) > 4) And (mMouseOverTabIndex > 0)
         If BeginDrag Then 'OK, we're dragging. But which type of drag operation?
            mDragDistance = DragDistance
            If Shift And vbCtrlMask Then
               mScrollDragging = True
            Else
               mReOrderDragging = True
               mDraggedTabIndex = mMouseOverTabIndex
               mDragFromPosition = PositionForTabIndex(mDraggedTabIndex)
               DoDragReOrdering Sgn(mDragDistance)
            End If
            mMouseOverTabIndex = 0
         End If
      Else 'we're already in the process of a drag op
         If mScrollDragging Then
            ShiftViewport mDragDistance - DragDistance
         ElseIf mReOrderDragging Then 'let the timer take over
            Tabs(mDraggedTabIndex).x = Tabs(mDraggedTabIndex).HomeX + x - mMouseDownX
            mDragDirection = Sgn(DragDistance - mDragDistance)
            mDragDistance = DragDistance
            Select Case True
               Case Tabs(mDraggedTabIndex).x <= mViewportX 'before the beginning of the visible area
                  Tabs(mDraggedTabIndex).x = mViewportX
                  mAutoScrollDirection = -1
               Case Tabs(mDraggedTabIndex).x + Tabs(mDraggedTabIndex).Width >= mViewportW + mViewportX 'beyond the end of the visible area
                  Tabs(mDraggedTabIndex).x = mViewportX + mViewportW - Tabs(mDraggedTabIndex).Width
                  mAutoScrollDirection = 1
                  If Tabs(mDraggedTabIndex).x + Tabs(mDraggedTabIndex).Width > mTotalTabsW Then 'beyond the end of the last tab
                     Tabs(mDraggedTabIndex).x = mTotalTabsW - Tabs(mDraggedTabIndex).Width
                  End If
               Case Else
                  If Tabs(mDraggedTabIndex).x + Tabs(mDraggedTabIndex).Width > mTotalTabsW Then 'beyond the end of the last tab
                     Tabs(mDraggedTabIndex).x = mTotalTabsW - Tabs(mDraggedTabIndex).Width
                  End If
                  mAutoScrollDirection = 0
                  DoDragReOrdering mDragDirection
            End Select
         End If
      End If
   End If

End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mMouseDownX = -1
   mMouseUpButton = Button
   
   If Button <> vbLeftButton Then Exit Sub
   
   If mReOrderDragging Then
      DoDragEnd
   ElseIf mScrollDragging Then
      mScrollDragging = False
   End If
   Refresh
End Sub
Private Sub DoDragEnd()
   
   mReOrderDragging = False
   mMouseDownX = -1
   tMouseTimer.Interval = 40
   
   If mDragToPosition = 1 Then
      Tabs(mDraggedTabIndex).HomeX = 0
   Else
      Tabs(mDraggedTabIndex).HomeX = Tabs(mPositions(mDragToPosition - 1)).HomeX + Tabs(mPositions(mDragToPosition - 1)).Width
   End If

   mDraggedTabIndex = 0

   Do Until ParkTabs(0, 8)
      Sleep 20
   Loop

End Sub
Private Function CloseButtonHitTest(ByVal x As Single, ByVal y As Single) As Boolean
Dim R As RECT

   With Tabs(mMouseOverTabIndex)
      R.Left = .x + .Width - 19
      R.Top = 6
      R.Right = R.Left + 13
      R.Bottom = TAB_H
   End With
   CloseButtonHitTest = CBool(PtInRect(R, x, y))

End Function
Private Function HitTest(x As Single) As Long
Dim i As Long
   For i = 1 To UBound(Tabs)
      If (x >= Tabs(i).x) And (x <= Tabs(i).x + Tabs(i).Width) Then
         HitTest = i
         Exit Function
      End If
   Next i
End Function
Private Function PositionForTabIndex(Index As Long) As Long
Dim i As Long
   For i = 1 To mPositions.Count
      If mPositions(i) = Index Then Exit For
   Next i
   If i <= mPositions.Count Then PositionForTabIndex = i
End Function
Public Sub Refresh()
Dim i As Long
   If mLockUpdate Then Exit Sub
   
   With UserControl
      .AutoRedraw = True
      .Cls
      For i = 1 To UBound(Tabs)
         If mDraggedTabIndex <> i Then
            If Not Tabs(i).Hidden Then DrawTab Tabs(i), mActiveTabIndex = i, mMouseOverTabIndex = i, mDraggedTabIndex = i, mCloseButtonIsHot
         End If
      Next i
      If mReOrderDragging Then DrawTab Tabs(mDraggedTabIndex), mActiveTabIndex = mDraggedTabIndex, True, True, mCloseButtonIsHot
      If mButtonWidth > 0 Then DrawScrollButtons 1: DrawScrollButtons 2
      .AutoRedraw = False
   End With
End Sub
Private Sub DrawTab(thisTab As TabItem, IsActiveTab As Boolean, HasMouseOver As Boolean, IsDraggedTab As Boolean, IsOverCloseButton As Boolean)
Dim R As RECT, clr As Long
   'the whole drawing thing could be made a lot nicer with GDI, but this is fine, for now
   R.Left = thisTab.x - mViewportX + mButtonWidth + 2
   R.Top = 3
   R.Right = R.Left + thisTab.Width - 4
   R.Bottom = TAB_H - 3
   
   If R.Right < 0 Or R.Left > UserControl.ScaleWidth Then Exit Sub
   
   With UserControl
      If IsActiveTab Then
         clr = vbWhite
      Else
         If HasMouseOver Then clr = RGB(230, 230, 230) Else clr = RGB(220, 220, 220)
      End If
      
      .DrawWidth = 1
      .ForeColor = IIf(IsActiveTab, vbBlack, RGB(100, 100, 100))
      If mScrollDragging Or mReOrderDragging Then .ForeColor = IIf(IsDraggedTab, vbBlack, RGB(180, 180, 180))
      .FillColor = clr
      Rectangle .hdc, R.Left, R.Top, R.Right, R.Bottom
      
      If HasMouseOver And Not IsDraggedTab Then
         .DrawWidth = 2
         .ForeColor = IIf(IsOverCloseButton, vbRed, RGB(160, 160, 160))
         .FillColor = IIf(IsOverCloseButton, vbRed, RGB(160, 160, 160))
         Rectangle .hdc, R.Right - 15, R.Top + 5, R.Right - 4, R.Bottom - 4
         
         .ForeColor = vbWhite
         MoveToEx .hdc, R.Right - 14, 9, 0&
         LineTo .hdc, R.Right - 7, TAB_H - 10
         MoveToEx .hdc, R.Right - 14, TAB_H - 10, 0&
         LineTo .hdc, R.Right - 7, 9
      End If
      
      If thisTab.IconHandle <> 0 Then DrawIconEx .hdc, R.Left + 2, 5, thisTab.IconHandle, 16, 16, 0, 0, DI_NORMAL
      
      .ForeColor = vbBlack
      R.Left = R.Left + ICON_SPACE_W + 2
      R.Right = R.Right - IIf(HasMouseOver And Not IsDraggedTab, 16, 0)
      DrawText .hdc, thisTab.Caption, -1, R, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
      
   End With

End Sub
Private Sub DrawScrollButtons(Index As Long)
Dim L As Long, R As Long, Sign As Long, clr As Long

   With UserControl
      If Index = 1 Then
         L = BTTN_W: R = L - BTTN_W: Sign = 1
         clr = IIf(mHotButtonIndex = -Index, vbButtonText, IIf(mViewportX > 0, vbButtonText, vbScrollBars))
         .FillColor = IIf(Abs(mHotButtonIndex) = Index And mViewportX > 0, RGB(220, 220, 220), .BackColor)
      Else
         L = .ScaleWidth - BTTN_W: R = L + BTTN_W: Sign = -1
         clr = IIf(mHotButtonIndex = -Index, vbButtonText, IIf(mViewportX < mMaxViewportX, vbButtonText, vbScrollBars))
         .FillColor = IIf(Abs(mHotButtonIndex) = Index And mViewportX < mMaxViewportX, RGB(220, 220, 220), .BackColor)
      End If
   
      .DrawWidth = 1
      .ForeColor = RGB(180, 180, 180)
      Rectangle .hdc, L, -1, R - 1 * Sign, TAB_H + 1
      
      .DrawWidth = 3
      .ForeColor = clr
      MoveToEx .hdc, L - 8 * Sign, 8, 0&
      LineTo .hdc, R + 6 * Sign, TAB_H / 2
      LineTo .hdc, L - 8 * Sign, TAB_H - 8
   End With
End Sub
Private Function ShiftViewport(ShiftAmount As Long, Optional SkipRefresh As Boolean) As Long
Dim NewX As Long
   NewX = mViewportX + ShiftAmount
   If NewX < 0 Then NewX = 0
   If NewX > mMaxViewportX Then NewX = mMaxViewportX
   ShiftViewport = NewX - mViewportX 'return the amount we shifted
   mViewportX = NewX
   If Not SkipRefresh Then Refresh
End Function
Private Sub tMouseTimer_Timer() 'the timer can be doing nothing, or either of [A], [B] or [C], as described below
Dim pt As POINTAPI, H As Long, ShiftedAmount As Long
   
   If mHotButtonIndex < 0 Then 'case [A], the timer's job is to scroll the control (user mouse-down on one of the buttons)
      If mHotButtonIndex = -1 Then ShiftViewport -10 Else ShiftViewport 10
      If mHotButtonIndex = -1 And mViewportX <= 0 Then mHotButtonIndex = 0
      If mHotButtonIndex = -2 And mViewportX >= mMaxViewportX Then mHotButtonIndex = 0
      Refresh
   ElseIf mReOrderDragging Then 'case [B], the timer's job is to keep up with the user's dragging around of the tabs
      tMouseTimer.Interval = 25 ' we'll bump-up the frequency for that
      
      H = GetCapture
      If H <> UserControl.hwnd Then
         DoDragEnd 'check that we still have mouse-capture, and end the drag-op if we don't
      Else
         If mAutoScrollDirection <> 0 Then
            ShiftedAmount = ShiftViewport(mAutoScrollDirection * 5, True)
            Tabs(mDraggedTabIndex).x = Tabs(mDraggedTabIndex).x + ShiftedAmount
            DoDragReOrdering mAutoScrollDirection
         End If
         
         ParkTabs mDraggedTabIndex 'send any tabs that have been displaced by the drag-op, to their new homes
      End If
   Else 'otherwise, [C] the timer is used to monitor mouse entry and exit
      If Not tMouseTimer.Enabled Then  '[C1] a mouse-enter, effectively...
         tMouseTimer.Interval = 40
         mMouseOverTabIndex = 0
         tMouseTimer.Enabled = True '... so enable the timer
      Else '[C2] we're testing for a mouse-leave
         GetCursorPos pt
         If Not WindowFromPoint(pt.x, pt.y) = UserControl.hwnd Then
            If mHotButtonIndex > 0 Then mHotButtonIndex = 0
            mMouseOverTabIndex = 0
            tMouseTimer.Enabled = False '... and if it happened, we disable the timer
            Refresh '...and refresh the control
         End If
      End If
   End If

End Sub
Private Sub UserControl_Initialize()
   lblScroll(1).Width = BTTN_W: lblScroll(2).Width = BTTN_W
   lblScroll(1).BorderStyle = 0: lblScroll(2).BorderStyle = 0
   mMouseDownX = -1
   UserControl.FillStyle = 0
   UserControl.Font = "Tahoma"
   UserControl.FontSize = 8
   Set mPositions = New Collection
   ReDim Tabs(0)
End Sub
Private Sub UserControl_Resize()
   ReCalcMetrics
   If mActiveTabIndex > 0 Then
      mLockUpdate = True
      EnsureTabVisible Tabs(mPositions(mActiveTabIndex)).Key
      mLockUpdate = False
   End If
   Refresh
End Sub
Private Sub ReCalcMetrics()
   lblScroll(1).Move 0, 0, BTTN_W: lblScroll(2).Move UserControl.ScaleWidth - lblScroll(2).Width, 0, BTTN_W
   If mTotalTabsW <= UserControl.ScaleWidth Then mButtonWidth = 0 Else mButtonWidth = BTTN_W
   lblScroll(1).Visible = (mButtonWidth > 0): lblScroll(2).Visible = (mButtonWidth > 0)
   mViewportW = UserControl.ScaleWidth - (2 * mButtonWidth)
   mMaxViewportX = mTotalTabsW - mViewportW
   If mMaxViewportX < 0 Then mMaxViewportX = 0
End Sub
Private Sub UserControl_Terminate()
   Set mPositions = Nothing
End Sub

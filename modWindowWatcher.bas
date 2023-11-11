Attribute VB_Name = "modWindowWatcher"
Option Explicit

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private mDict As Dictionary
Public Function EnumIDEWindows() As Dictionary
   Set mDict = New Dictionary
   If gIsSDI Then
      EnumThreadWindows App.ThreadID, AddressOf EnumSDIWindowsProc, 0&
   Else
      EnumChildWindows gMDIhWnd, AddressOf EnumMDIWindowsProc, ByVal 0&
   End If
   Set EnumIDEWindows = mDict
End Function
Private Function EnumMDIWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim ParentHwnd As Long, sClassName As String
   ParentHwnd = GetParent(hWnd)
   If ParentHwnd = gMDIhWnd Then
      sClassName = ClassName(hWnd)
      If sClassName <> "VBMdiChildHack" Then mDict.Add hWnd, sClassName
   End If
   EnumMDIWindowsProc = 1
End Function
Private Function EnumSDIWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim ParentHwnd As Long, sClassName As String
   ParentHwnd = GetWindow(hWnd, GW_OWNER)
   If ClassName(ParentHwnd) = "IDEOwner" Then
      sClassName = ClassName(hWnd)
      Select Case sClassName
         Case "VbaWindow", "DesignerWindow", "ThunderForm", "ThunderMDIForm", "ThunderDFrame"
            mDict.Add hWnd, sClassName
      End Select
   End If
   EnumSDIWindowsProc = 1
End Function
Public Function CaptionForHwnd(hWnd As Long) As String
Dim Title As String * 255, tLen As Long
   tLen = GetWindowTextLength(hWnd)
   GetWindowText hWnd, Title, 255
   CaptionForHwnd = Left$(Title, tLen)
End Function
Public Function ClassName(hWnd As Long) As String
Dim buf As String, buflen As Long
   buflen = 256
   buf = Space$(buflen - 1)
   buflen = GetClassName(hWnd, buf, buflen)
   buf = Left$(buf, buflen)
   ClassName = buf
End Function

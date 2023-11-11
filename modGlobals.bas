Attribute VB_Name = "modGlobals"
Option Explicit

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const TOOLBARNAME = "ceTabStrip"
Public Const TOOLBAR_HEIGHT As Long = 26

Public gLaunchedAtIDEStartup  As Boolean 'indicates that Add-In launched at IDE start-up
Public gGuiForm               As frmTabStrip
Public gOptionsForm           As frmTabStripOptions
Public gVBInstance            As VBIDE.VBE
Public gIsSDI                 As Boolean
Public gShowFullWidth         As Boolean
Public gMaintainVBW           As Boolean
Public gAlwaysMaximise        As Boolean
Public gEnsureVisible         As Boolean
Public gLogging               As Boolean
Public gIDEhWnd               As Long  'handle to the Main IDE window
Public gMDIhWnd               As Long  'handle to the MDI window
Public gToolbarHwnd           As Long  'handle to the Office Toolbar (Const TOOLBARNAME) which hosts our UserControl
Public Sub Log(MsgText As String)
   gOptionsForm.Log MsgText
End Sub
Public Sub InitGlobalIDEVariables()
Dim hWnd As Long
   gIDEhWnd = gVBInstance.MainWindow.hWnd
   gMDIhWnd = FindWindowEx(gIDEhWnd, 0&, "MDIClient", vbNullString)
   hWnd = FindWindowEx(gIDEhWnd, 0&, "MsoCommandBarDock", "MsoDockTop")
   gToolbarHwnd = FindWindowEx(hWnd, 0&, "MsoCommandBar", TOOLBARNAME)
End Sub


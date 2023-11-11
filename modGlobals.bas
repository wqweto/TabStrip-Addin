Attribute VB_Name = "modGlobals"
Option Explicit

'--- for LOGFONT
Private Const FW_NORMAL                     As Long = 400
Private Const LF_FACESIZE                   As Long = 32
'--- for SystemParametersInfo
Private Const SPI_GETICONTITLELOGFONT       As Long = 31
'--- for GetDeviceCaps
Private Const LOGPIXELSY                    As Long = 90        '  Logical pixels/inch in Y

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpStr As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Type LOGFONTW
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(0 To LF_FACESIZE - 1) As Integer
End Type

Public Const TOOLBARNAME = "ceTabStrip"

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
Public gToolbarHeight         As Long
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

Property Get SystemIconFont() As StdFont
    Dim uFont           As LOGFONTW
    Dim sBuffer         As String
    Dim lLogPixels      As Long
    
    Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, LenB(uFont), uFont, 0)
    Set SystemIconFont = New StdFont
    With SystemIconFont
        sBuffer = Space$(lstrlen(VarPtr(uFont.lfFaceName(0))))
        Call CopyMemory(ByVal StrPtr(sBuffer), uFont.lfFaceName(0), LenB(sBuffer))
        .Name = sBuffer
        .Bold = (uFont.lfWeight > FW_NORMAL)
        .Charset = uFont.lfCharSet
        .Italic = (uFont.lfItalic <> 0)
        .Strikethrough = (uFont.lfStrikeOut <> 0)
        .Underline = (uFont.lfUnderline <> 0)
        .Weight = uFont.lfWeight
        lLogPixels = pvGetLogPixels()
        If lLogPixels <> 0 Then
            .Size = -(uFont.lfHeight * 72#) / lLogPixels
        End If
    End With
End Property

Private Function pvGetLogPixels() As Long
    Dim hTempDC         As Long
    
    hTempDC = GetDC(0)
    pvGetLogPixels = GetDeviceCaps(hTempDC, LOGPIXELSY)
    Call ReleaseDC(0, hTempDC)
End Function


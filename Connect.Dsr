VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10140
   ClientLeft      =   10950
   ClientTop       =   3690
   ClientWidth     =   23700
   _ExtentX        =   41804
   _ExtentY        =   17886
   _Version        =   393216
   Description     =   "Adds a Tab-Strip, with one tab per open code module"
   DisplayName     =   "Tab-Strip AddIn"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Tab-Strip Add-In"
Option Explicit

Public GUIvisible             As Boolean

Dim mcbMenuCommandBar         As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Public Sub Hide()
   On Error Resume Next
   GUIvisible = False
End Sub
Public Sub Show()
   On Error GoTo Show_Error

   If GUIvisible Then Exit Sub
   
   If gOptionsForm Is Nothing Then Set gOptionsForm = New frmTabStripOptions
   Load gOptionsForm
   If gLogging Then Log "Addin-In Started with the IDE?: " & gLaunchedAtIDEStartup
   
   CreateToolbar
   
   InitGlobalIDEVariables
   
   If gGuiForm Is Nothing Then Set gGuiForm = New frmTabStrip
   Load gGuiForm
   gGuiForm.Init
   
   GUIvisible = True
   
   If gLogging Then Log "UI is now visible - " & IIf(gIsSDI, "SDI", "MDI") & " mode"

   Exit Sub

Show_Error:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Show of Connect"

End Sub
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
   On Error GoTo error_handler
    
   Set gVBInstance = Application
    
   gIsSDI = (gVBInstance.DisplayModel = vbext_dm_SDI)
    
   If ConnectMode = ext_cm_External Then
      Me.Show
   Else 'ext_cm_Startup=1 ; ext_cm_AfterStartup=0
      Set mcbMenuCommandBar = AddToAddInCommandBar("Tab-Strip Options...")
      Set MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
   End If
  
   If ConnectMode = ext_cm_AfterStartup Then
      Me.Show
   Else
      gLaunchedAtIDEStartup = True
   End If
  
   Exit Sub
    
error_handler:
    MsgBox "AddinInstance_OnConnection: " & Err.Description
End Sub
Private Sub AddinInstance_OnStartupComplete(custom() As Variant) 'IDE up-and-running
   Me.Show
End Sub
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant) ' Add-In removed from VB
   On Error GoTo AddinInstance_OnDisconnection_Error

   If GUIvisible Then
       SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
       GUIvisible = False
   Else
       SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"  '"0"
   End If
    
   Set MenuHandler = Nothing
   mcbMenuCommandBar.Delete
   Set mcbMenuCommandBar = Nothing
    
   Unload gGuiForm
   Set gGuiForm = Nothing
   Unload gOptionsForm
   Set gOptionsForm = Nothing
   DeleteToolBar
   Set gVBInstance = Nothing

   Exit Sub

AddinInstance_OnDisconnection_Error:

   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in AddinInstance_OnDisconnection"
End Sub
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
   If GUIvisible Then gOptionsForm.Show
End Sub
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
Dim cbMenuCommandBar As Office.CommandBarControl
Dim cbMenu As Office.CommandBar
  
   On Error GoTo ErrHandler
   
   Set cbMenu = gVBInstance.CommandBars("Add-Ins")
   If cbMenu Is Nothing Then Exit Function
   
   Set cbMenuCommandBar = cbMenu.Controls.Add(1)
   cbMenuCommandBar.Caption = sCaption
   
   Set AddToAddInCommandBar = cbMenuCommandBar
   
   Exit Function
    
ErrHandler:

End Function
Private Sub CreateToolbar()
Dim i As Long

   On Error GoTo CreateToolbar_Error

   DeleteToolBar
   gVBInstance.CommandBars.Add TOOLBARNAME, msoBarTop, , True

   For i = 0 To Screen.Width / Screen.TwipsPerPixelX / 550
      AddButton '- the idea is to add as many of these as is necessary to make the toolbar a full screen width. Big kludge!
   Next i

   With gVBInstance.CommandBars(TOOLBARNAME)
      .Visible = True
      .Protection = msoBarNoMove
      .Height = TOOLBAR_HEIGHT
   End With

   On Error GoTo 0
   Exit Sub

CreateToolbar_Error:
   If Err.Number = -2147467259 And InStr(Err.Description, "'~'") > 0 Then 'Grrrrr. The infamous Method '~ of Object '~' error.
      Resume Next
   Else
      MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateToolbar of Designer Connect"
   End If

End Sub
Private Sub AddButton()
   On Error GoTo ErrHandler
   With gVBInstance.CommandBars(TOOLBARNAME).Controls.Add
      .Height = TOOLBAR_HEIGHT - 2 'seems to add 2 border pixels
      .Style = msoButtonCaption
      .Width = 550 'appears to be the max
      .Enabled = False
   End With
ErrHandler:
End Sub
Private Sub DeleteToolBar()
   On Error Resume Next
   gVBInstance.CommandBars(TOOLBARNAME).Delete
End Sub

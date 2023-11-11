Attribute VB_Name = "modWorkspaceVBW"
' Creates a VBW file file of all open code panes/designers, in the order that their corresponding tabs appear on the TabStrip
' Due to the need to preserve this order, some component entries will be duplicated - but the IDE doesn't seem to care about that
' e.g. where a code pane is on Tab 1 and its designer is on Tab4, you would see ComponentA, ComponentB, ComponentC, ComponentA
' This add-in will later use the data from the VBW file to restore the correct order when a project is re-opened.
' Note that this 'workspace preservation' only occurs upon a project being saved! This is because, upon exit, some windows (i.e. designers)
' are destroyed before this sub would get an opportunity to see them.
Option Explicit
Const DESIGNER_FLAG As String = "*"
Private Declare Function GetMem4 Lib "msvbvm60" (src As Any, Dst As Any) As Long
Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwFileAttributes As VbFileAttribute) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private mWorkspace As Dictionary, mComponentHwnds As Dictionary
Public Sub SaveVBWFile(Project As VBProject)  'Only used for MDI layouts
Dim CP As CodePane, VBComp As VBComponent, VBProj As VBProject, hWnd As Long

   On Error GoTo ErrHandler
   
   If gLogging Then Log "Saving workspace for " & Project.Name
   
   CreateWorkspaceInfo Project, False

   WriteVBWFile Project.FileName
   
   Set mWorkspace = Nothing
   Set mComponentHwnds = Nothing
   
   Exit Sub
   
ErrHandler:
   If gLogging Then Log "SaveVBWFile: " & Err.Number & " - " & Err.Description
End Sub
Private Sub WriteVBWFile(ByVal FilePath As String)
Dim FileNo As Integer, i As Long, TabKey As String
   
   On Error GoTo ErrHandler
   
   FilePath = Replace(FilePath, ".vbp", ".vbw")
   SetFileAttributesW StrPtr(FilePath), vbNormal
   DeleteFileW (StrPtr(FilePath))
   
   FileNo = FreeFile
   
   Open FilePath For Output As #FileNo
      For i = 1 To gGuiForm.ucTabStrip.TabCount
         TabKey = gGuiForm.ucTabStrip.KeyFromPosition(i)
         If mWorkspace.Exists(mComponentHwnds(TabKey)) Then
            Print #FileNo, mWorkspace(mComponentHwnds(TabKey))
         End If
      Next i
   Close #FileNo
   
   SetFileAttributesW StrPtr(FilePath), vbReadOnly
   
   Exit Sub

ErrHandler:
   If gLogging Then Log "WriteVBWFile: " & Err.Number & " - " & Err.Description
End Sub
Public Sub RestoreFromVBW(Project As VBProject)
Dim FilePath As String, FileNo As Integer, VBWInfo As Collection, sLine As String, sComp As String, s() As String, i As Long

   On Error GoTo ErrHandler

   If gLogging Then Log "Restoring Tab order from VBW file for " & Project.Name

   Set VBWInfo = New Collection
   
   FilePath = Replace(Project.FileName, ".vbp", ".vbw")

   FileNo = FreeFile
   
   Open FilePath For Input As #FileNo
      Do While Not EOF(FileNo)
         Line Input #FileNo, sLine
         s = Split(sLine, "=")
         sComp = Trim$(s(0))
         If InStr(s(1), "Z") = 0 Then
            sComp = sComp & DESIGNER_FLAG
         End If
         VBWInfo.Add sComp, sComp
      Loop
   Close #FileNo
   
   CreateWorkspaceInfo Project, True
   
   For i = 1 To VBWInfo.Count
      'If gLogging Then Log VBWInfo(i) & " found in VBW file"
      If mComponentHwnds.Exists(VBWInfo(i)) Then
         'Log "..." & VBWInfo(i) & " has a corresponding window (" & mComponentHwnds(VBWInfo(i)) & ")"
         If gGuiForm.ucTabStrip.TabExists(mComponentHwnds(VBWInfo(i))) Then
            'Log "......" & VBWInfo(i) & " also has a corresponding tab"
            gGuiForm.ucTabStrip.MoveTab mComponentHwnds(VBWInfo(i))
         'Else
         '   Log "......" & VBWInfo(i) & "has no corresponding tab"
         End If
      'Else
         'Log "..." & VBWInfo(i) & " has no corresponding window"
      End If
   Next i
   
   gGuiForm.ucTabStrip.Init
   
   Set mWorkspace = Nothing
   Set mComponentHwnds = Nothing
   
   Exit Sub
ErrHandler:
   If gLogging Then Log "RestoreFromVBW: " & Err.Number & " - " & Err.Description
   
End Sub
Private Sub CreateWorkspaceInfo(Project As VBProject, IsRestoreMode As Boolean) 'won't work if we've hit F5! (So we don't ever call it in that situation)
Dim CP As CodePane, VBComp As VBComponent, hWnd As Long, sVBWInfo As String

   On Error GoTo ErrHandler
   
   Set mWorkspace = New Dictionary 'Key: ComponentName, Value: Workspace data
   Set mComponentHwnds = New Dictionary 'Key: hWnd, Value: ComponentName (or vice verse if restoring)
   
   For Each CP In gVBInstance.CodePanes
      Set VBComp = CP.CodeModule.Parent
      If VBComp.Collection.Parent = Project.Name Then
         GetMem4 ByVal ObjPtr(CP.Window) + &H1C, hWnd
         If IsRestoreMode Then
            mComponentHwnds.Add VBComp.Name, "W" & hWnd
         Else
            mComponentHwnds.Add "W" & hWnd, VBComp.Name
         End If
         With CP.Window
            sVBWInfo = VBComp.Name & " = " & .Left & "," & .Top & "," & .Width & "," & .Height & ",Z,"
         End With
         If VBComp.HasOpenDesigner Then
            GetMem4 ByVal ObjPtr(VBComp.DesignerWindow) + &H1C, hWnd
            If IsRestoreMode Then
               mComponentHwnds.Add VBComp.Name & DESIGNER_FLAG, "W" & hWnd
            Else
               mComponentHwnds.Add "W" & hWnd, VBComp.Name & DESIGNER_FLAG
            End If
            With VBComp.DesignerWindow
               sVBWInfo = sVBWInfo & .Left & "," & .Top & "," & .Width & "," & .Height & ","
            End With
         Else
            sVBWInfo = sVBWInfo & "0,0,0,0,C"
         End If
         mWorkspace.Add VBComp.Name, sVBWInfo
      End If
   Next CP
   
   For Each VBComp In Project.VBComponents
      If VBComp.HasOpenDesigner Then
         If Not mWorkspace.Exists(VBComp.Name) Then
            GetMem4 ByVal ObjPtr(VBComp.DesignerWindow) + &H1C, hWnd
            With VBComp.DesignerWindow
               If IsRestoreMode Then
                  mComponentHwnds.Add VBComp.Name & DESIGNER_FLAG, "W" & hWnd
               Else
                  mComponentHwnds.Add "W" & hWnd, VBComp.Name & DESIGNER_FLAG
               End If
               mWorkspace.Add VBComp.Name & DESIGNER_FLAG, VBComp.Name & " = " & "0,0,0,0,C," & .Left & "," & .Top & "," & .Width & "," & .Height & ","
            End With
         Else
            mWorkspace.Add VBComp.Name & DESIGNER_FLAG, Replace(mWorkspace(VBComp.Name), "Z", " ")
         End If
      End If
   Next VBComp
   
   Exit Sub
   
ErrHandler:
   If gLogging Then Log "CreateWorkspaceInfo: " & Err.Number & " - " & Err.Description

End Sub


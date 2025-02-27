
Sub ExportMacroToGitHub()
    Dim fso As Object
    Dim File As Object
    Dim GitFile As Object
    Dim ModuleCode As String
    Dim GitHubFolder As String
    
    ' Set GitHub Folder Path
    GitHubFolder = "C:\Users\grzyb\Desktop\react-website-v3-main\public\files"

    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create Folder if it doesn't exist
    If fso.FolderExists(GitHubFolder) = False Then
        fso.CreateFolder GitHubFolder
    End If
    
    ' Loop through VBA Components
    Dim Component As Object
    
    For Each Component In Application.VBE.VBProjects(1).VBComponents
        If Component.Type = 1 Then ' 1 = Module
            ModuleCode = Component.CodeModule.Lines(1, Component.CodeModule.CountOfLines)
            
            ' Check if Module Code is not empty
            If ModuleCode <> "" Then
                Set GitFile = fso.CreateTextFile(GitHubFolder & "\" & Component.Name & ".bas", True)
                GitFile.Write ModuleCode
                GitFile.Close
            End If
        End If
    Next Component
    
    MsgBox "Export Successful!"
    
    ' Automate Git Commit & Push
    Dim Path As String
    Path = "cd /d " & GitHubFolder & " && git add . && git commit -m ""VBA Macros Exported"" && git push"
    Shell "cmd.exe /c " & Path, vbNormalFocus
End Sub

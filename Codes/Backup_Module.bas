Attribute VB_Name = "Backup_Module"
Option Compare Database

    Sub RunPythonScript()
    Dim srcfile As String
    Dim backfile As String
    Dim backfolder As String
    Dim fso As Object
    Dim timestamp As String
    
    srcfile = Application.CurrentDb.Name
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    backfolder = Application.CurrentProject.Path & "\Backup"
        If Not fso.FolderExists(backfolder) Then
            fso.CreateFolder (backfolder)
        End If
        
    timestamp = Format(Now(), "yyyy_mm_dd_hh_mm_ss")
    
    backfile = backfolder & "\Backup_" & timestamp & ".accdb"
    
    fso.CopyFile srcfile, backfile, True
    
    Set fso = Nothing
    End Sub


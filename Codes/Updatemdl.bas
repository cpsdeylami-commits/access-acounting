Attribute VB_Name = "Updatemdl"
Option Compare Database

Sub UpdateVersion()
    Dim srcDB As DAO.Database
    Dim destDB As DAO.Database
    Dim tableNames As Variant
    Dim i As Integer
    
    tableNames = Array("cstmr", "etc", "indetail", "INVOICE", "jari", "rprt_tbl", "Sale", "Store")
    
    Dim srcPath As String
    Dim destPath As String
    srcPath = Application.CurrentProject.Path & "\src.accdb"
    destPath = Application.CurrentProject.Path & "\OC.accdb"
    
    If IsNull(srcPath) Then
    MsgBox "Source File Not Found !", vbInformation
    Exit Sub
    Else
    Set srcDB = OpenDatabase(srcPath)
    Set destDB = OpenDatabase(destPath)
    
    
    For i = LBound(tableNames) To UBound(tableNames)
        Dim srcTable As String
        Dim destTable As String
        
        srcTable = tableNames(i)
        destTable = tableNames(i)
        
        destDB.Execute "DELETE FROM " & destTable, dbFailOnError
        
        destDB.Execute "INSERT INTO " & destTable & " SELECT * FROM [MS Access;DATABASE=" & srcPath & "]." & srcTable, dbFailOnError
    Next i
    
    
    srcDB.Close
    destDB.Close
    Set srcDB = Nothing
    Set destDB = Nothing
    
    Dim stru As String
    stru = "INSERT INTO updtvrsn (versionupdate) VALUES ('" & 10 & "')"
    DoCmd.RunSQL stru

    MsgBox "Data copied successfully for all tables!", vbInformation
    Form_popupform.Command113.Enabled = False
    End If
End Sub

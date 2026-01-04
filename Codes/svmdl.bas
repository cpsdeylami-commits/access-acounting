Attribute VB_Name = "svmdl"
Option Compare Database

Public Sub svlu()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim s_v As Double
    
    Set db = CurrentDb()
    
    s_v = 0
    
    Set rs = db.OpenRecordset("SELECT ValueOA FROM Store")
    
    Do While Not rs.EOF
        s_v = s_v + rs!ValueOA
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
        
    Form_Reprt_frm.Label14.Caption = s_v
End Sub


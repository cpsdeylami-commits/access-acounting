Attribute VB_Name = "tsmdl"
Option Compare Database

Public Sub totlstr()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim t_s As Double
    
    Set db = CurrentDb()
    
    t_s = 0
    
    Set rs = db.OpenRecordset("SELECT Quntity FROM Store")
    
    Do While Not rs.EOF
        t_s = t_s + rs!Quntity
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
        
    Form_Reprt_frm.Label13.Caption = t_s
End Sub


Attribute VB_Name = "asmdl"
Option Compare Database

Public Sub asupdt()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim a_s As Double
    
    Set db = CurrentDb()
    
    a_s = 0
    
    Set rs = db.OpenRecordset("SELECT OutQuantity FROM Sale")
    
    Do While Not rs.EOF
        a_s = a_s + rs!OutQuantity
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
        
    Form_Reprt_frm.Label11.Caption = a_s
End Sub

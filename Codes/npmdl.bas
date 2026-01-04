Attribute VB_Name = "npmdl"
Option Compare Database

Public Sub npupdt()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim n_p As Double
    
    Set db = CurrentDb()
    
    n_p = 0
    
    Set rs = db.OpenRecordset("SELECT * FROM rprt_tbl")
    
    Do While Not rs.EOF
        If IsNull(rs!return) Then
           n_p = n_p + rs!netprft
        Else
        n_p = n_p + rs!netprft
        n_p = n_p - (rs!return * -1)
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
        
    Form_Reprt_frm.Label12.Caption = n_p
End Sub


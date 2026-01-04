VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Salesrchqry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command13_Click()
End If

End Sub

Private Sub Command11_Click()
Dim rtr1 As Long
Dim rtr2 As Long
Dim srtr As Double
Dim ss As Double

rtr2 = Me.Ourprice.Value

Dim respond As Integer
respond = MsgBox("Do you Really Want to Delete This Recoed ?", vbYesNo + vbQuestion, "Delete Notification")
    
    If respond = vbYes Then
    
    Dim db As DAO.Database
    Dim rst As String
    Dim rsstor As String
        
    Dim q As Integer
    Dim c As String
    c = Me.IDs.Value
    q = CInt(Me.OutQuantity.Value)
    
    Dim dbu As DAO.Database
    Dim rsstr As DAO.Recordset
    Set dbu = CurrentDb
    Set rsstr = dbu.OpenRecordset("SELECT * FROM Store WHERE ID='" & c & "'")
    rtr1 = rsstr!inprice
    If Not rsstr.EOF Then
    rsstr.Edit
    rsstr!Quntity = rsstr!Quntity + q
    rsstr!ValueOA = Nz(rsstr!inprice, 0) * Nz(rsstr!Quntity, 0)
    rsstr.update
    End If
    rsstr.Close
    dbu.Close
    Set rsstr = Nothing
    Set dbu = Nothing
    Me.Dirty = False
    
    Set db = CurrentDb
    Dim u As Long
    u = CLng(Me.unictxt.Value)
    rst = "DELETE FROM Sale WHERE unic = " & u & " "
    db.Execute rst
    MsgBox "Data Was Deleted!", vbInformation
    Form_Salesrchqry.Refresh
    
    srtr = rtr1 - rtr2
    ss = srtr * q
    Dim strtr As String
    strtr = "INSERT INTO rprt_tbl (return) VALUES ('" & ss & "')"
    DoCmd.RunSQL strtr

    
    Else
    MsgBox "Opration Aborted !", vbInformation
    End If
    Exit Sub

End Sub

Private Sub Command12_Click()

End Sub

Private Sub Detail_Click()

End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500

End Sub

Private Sub IDs_Click()

End Sub

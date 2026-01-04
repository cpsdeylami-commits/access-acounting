VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Querysrchnam_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command11_Click()

Dim respond As Integer
respond = MsgBox("Do you Really Want to Delete This Recoed ?", vbYesNo + vbQuestion, "Delete Notification")
    
    If respond = vbYes Then
    
    Dim db As DAO.Database
    Dim rst As String

    Dim c As String
    c = Me.ID.Value

    Set db = CurrentDb
    rst = "DELETE FROM Store WHERE ID='" & c & "'"
    db.Execute rst
    MsgBox "Data Was Deleted!", vbInformation
    Form_Querysrchnam_frm.Refresh
    
    Else
    MsgBox "Opration Aborted !", vbInformation
    End If
    Exit Sub
End Sub

Public Sub Command13_Click()
Form_Querysrchnam_frm.Refresh

If IsNull(Me.ID.Value) Then
MsgBox "No ID Found !", vbInformation
Else

Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim i As Long
Dim s As Double
Dim j As Long
Dim c As String

c = Me.ID.Value
i = CLng(Me.Quntity.Value)
j = CLng(Me.inprice.Value)
s = i * j

    Set db = CurrentDb
    Set rst = db.OpenRecordset("SELECT * FROM Store WHERE ID='" & c & "'")
    If Not rst.EOF Then
    rst.Edit
    rst!ValueOA = Nz(rst!inprice, 0) * Nz(rst!Quntity, 0)
    rst.update
    End If
    rst.Close
    db.Close
    Set rst = Nothing
    Set db = Nothing
    Me.Dirty = False
    MsgBox "Data was Edited Successfuly!", vbInformation
End If
End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500

End Sub

Private Sub ID_Click()

End Sub

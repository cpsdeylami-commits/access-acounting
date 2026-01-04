VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Store_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim curnc As Integer

Private Sub Command18_Click()
Dim db As DAO.Database
Dim rst As DAO.Recordset


Dim strid As String
Dim strnam As String
Dim inqty As Long
Dim inprc As Long
Dim strdt As String

Dim strsql As String

If IsNull(Me.Text7) Or IsNull(Me.Text12) Or IsNull(Me.Text14) Or IsNull(Me.Text16) Or IsNull(Me.Text21) Or Me.Text7 = "" Or Me.Text12 = "" Or Me.Text16 = "" Or Me.Text14 = "" Or Me.Text21 = "" Then
    MsgBox "No Data Inserted !", vbInformation
Else
    strid = Me.Text7.Value
    strnam = Me.Text12.Value
    inqty = CLng(Me.Text14.Value)
    inprc = CLng(Me.Text16.Value) * curnc
    strdt = Me.Text21.Value

    strsql = "INSERT INTO Store (ID,PName,Quntity,inprice,Tarikh,ssn) VALUES ('" & strid & "','" & strnam & "','" & inqty & "','" & inprc & "','" & strdt & "','" & Text27.Value & "')"
    DoCmd.RunSQL strsql
    MsgBox "Add Data Successfull !", vbInformation
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("SELECT * FROM Store WHERE ID='" & strid & "'")
    If Not rst.EOF Then
    rst.Edit
    rst!ValueOA = Nz(rst!inprice, 0) * Nz(rst!Quntity, 0)
    rst.update
    End If
    rst.Close
    Set rst = Nothing
    Set db = Nothing
    Me.Dirty = False
    'êêMsgBox "ValueOA Updated!"
    
Me.Text7 = ""
Me.Text12 = ""
Me.Text14 = ""
Me.Text16 = ""

End If
Me.Text7.SetFocus
End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500


curnc = InputBox("Please Insert Today Currency(Yuan) Price :", "Currency Price")
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text21_Change()

End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text7_AfterUpdate()

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim stridchk As String
Dim newid As String

newid = Me.Text7.Value
Set db = CurrentDb

stridchk = "SELECT ID FROM Store WHERE ID='" & newid & "'"

Set rs = db.OpenRecordset(stridchk)

If Not rs.EOF Then
MsgBox "Your ID is Saved Before !", vbInformation
Me.Text7.Value = ""
Me.Text7.SetFocus
cancell = True
Me.Undo
End If

rs.Close
Set rs = Nothing
Set db = Nothing

End Sub

Private Sub Text7_BeforeUpdate(Cancel As Integer)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Text7_LostFocus()
End Sub

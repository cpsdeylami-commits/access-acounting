VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sale_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Combo30_Change()
Me.Text3.Value = DLookup("ID", "Store", "PName = Combo30.Text")

End Sub

Private Sub Combo30_GotFocus()
''Me.Combo30.ListIndex = DLookup("PName", "Store")

End Sub

Private Sub Command13_Click()
Dim db As DAO.Database
Dim rst As DAO.Recordset


Dim strid As String
Dim strnam As String
Dim oqty As Integer
Dim oprc As Long
Dim totl As Long

Dim strsqlout As String
Dim strsqlupdt As String
Dim nt As String



If IsNull(Me.Text3) Or IsNull(Me.Text5) Or IsNull(Me.Text7) Or IsNull(Me.Text9) Or IsNull(Me.Text11) Or Me.Text3 = "" Or Me.Text5 = "" Or Me.Text7 = "" Or Me.Text9 = "" Or Me.Text11 = "" Then
    MsgBox "No Data Inserted !", vbInformation
Else
    Dim t As Long
    Dim a As Long
    a = Me.Text3.Value
    t = Me.Text22.Value

    strid = Me.Text3.Value
    strnam = Me.Text5.Value
    oqty = CInt(Me.Text9.Value)
    oprc = CLng(Me.Text7.Value)
    totl = Me.Text11.Value
    
    Dim nc As Double
    Dim n1 As Long
    Dim m As Double
    Dim n2 As Long
    Dim n3 As Long
    n1 = Me.Text24.Value
    n2 = Me.Text11.Value
    n3 = Me.Text9.Value
    m = n1 * n3
    nc = n2 - m
    
    nt = "INSERT INTO rprt_tbl (netprft,ssn) VALUES ('" & nc & "','" & Text28.Value & "')"
    CurrentDb.Execute nt
    'MsgBox "Update 'nt,netc' Successfull !", vbInformation

    strsqlout = "INSERT INTO Sale (IDs,Poname,Ourprice,OutQuantity,TotalP,odate,netclc,ssn,byr) VALUES ('" & strid & "','" & strnam & "','" & oprc & "','" & oqty & "','" & totl & "','" & CLng(Me.Text26.Value) & "','" & m & "','" & Text28.Value & "','" & byrnam.Value & "')"
    DoCmd.RunSQL strsqlout
    MsgBox "Add Data Successfull !", vbInformation
    
    strsqlupdt = "UPDATE Store SET Quntity ='" & Me.Text22.Value & "' WHERE ID='" & a & "' "
    CurrentDb.Execute strsqlupdt
    'MsgBox "Update Data Quantity Successfull !", vbInformation

    
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
    'MsgBox "ValueOA Updated!"

Me.Text3 = ""
Me.Text5 = ""
Me.Text7 = ""
Me.Text9 = ""
Me.Text11 = ""

End If
Me.Text3.SetFocus
End Sub

Private Sub Text16_Change()
End Sub

Private Sub Command13_LostFocus()

End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500

End Sub

Private Sub Text11_AfterUpdate()

End Sub

Private Sub Text11_Change()
End Sub

Private Sub Text11_GotFocus()
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text3_AfterUpdate()
Me.Text5.Value = DLookup("PName", "Store", "ID=Text3")

Dim ip As Double
Dim sn As String
Dim op As Double
Dim ts As Double
ip = DLookup("inprice", "Store", "ID=Text3")
sn = DLookup("ssn", "Store", "ID=Text3")
op = (ip * 80) / 100
ts = ip + op
Me.Text7.Value = ts
Me.Text28.Value = sn
End Sub

Public Sub Text3_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Text3_LostFocus()

End Sub

Private Sub Text5_Exit(Cancel As Integer)

End Sub

Private Sub Text5_GotFocus()

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)
End Sub

Private Sub Text9_LostFocus()
If IsNull(Me.Text9.Value) Or IsNull(Me.Text7.Value) Then
MsgBox "Please Enter the value and quantity Correctly !", vbInformation
Exit Sub
Else
Form_Sale_frm.Refresh

Dim i As Integer
Dim j As Long
Dim s As Long
i = Me.Text9.Value
j = Me.Text7.Value
s = i * j
Me.Text11.Value = s

Dim t As Long
Dim a As Long
Dim a1 As Long
Dim b As Long
a = DLookup("Quntity", "Store", "ID=Text3")
a1 = DLookup("inprice", "Store", "ID=Text3")
If (i > a) Then
    MsgBox "Not enough Item available, AVAILABLE = " & a, vbInformation
    Me.Text9.Value = ""
    Form_Sale_frm.Text9.SetFocus
    Exit Sub
End If

b = Me.Text9.Value
t = a - b
Me.Text22.Value = CStr(t)
Me.Text24.Value = CStr(a1)
End If
End Sub

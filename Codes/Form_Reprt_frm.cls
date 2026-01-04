VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Reprt_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command30_Click()
Dim prft As Double
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rst As DAO.Recordset
Dim rst1 As DAO.Recordset
Dim rst2 As DAO.Recordset

Dim a As Double
Dim b As Double
Dim qsum As Double
Dim vsum As Double
Dim ntts As Double
Dim nts As Double

    If IsNull(Me.Text26.Value) Or IsNull(Me.Text28.Value) Then
    MsgBox "No Date Inserted !", vbInformation
    Else
    a = CDbl(Me.Text26.Value)
    b = CDbl(Me.Text28.Value)
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SUM(OutQuantity) AS ts FROM Sale WHERE odate BETWEEN " & a & " AND " & b & " AND ssn = '" & Me.Combo38.Value & "' ;")
    Set rst = db.OpenRecordset("SELECT SUM(TotalP) AS tv FROM Sale WHERE odate BETWEEN " & a & " AND " & b & " AND ssn = '" & Me.Combo38.Value & "' ;")
    Set rst1 = db.OpenRecordset("SELECT SUM(netclc) AS nc FROM Sale WHERE odate BETWEEN " & a & " AND " & b & " AND ssn = '" & Me.Combo38.Value & "' ;")
    Set rst2 = db.OpenRecordset("SELECT SUM(eql) AS eq FROM jari WHERE td BETWEEN " & a & " AND " & b & ";")

    If IsNull(rs!ts) Or IsNull(rst!tv) Or IsNull(rst1!nc) Or IsNull(rst2!eq) Then
    MsgBox "No Data Found !", vbInformation
    Exit Sub
    End If
    If Not rs.EOF Then
    qsum = qsum + rs!ts
    vsum = vsum + rst!tv
    ntts = rst!tv - rst1!nc
    nts = ntts - rst2!eq
    Else
    qsum = 0
    vsum = 0
    ntts = 0
    nts = 0
    End If
        
    ''outputs
    Me.Label11.Caption = Format(qsum, "#,###")
    Me.Label12.Caption = Format(vsum, "#,###")
    Me.Label21.Caption = Format(ntts, "#,###")
    Me.Label31.Caption = Format(nts, "#,###")
    Me.Text23.Value = rst1!nc
    ''calculate the profit in %''
    Dim sc1 As Double
    Dim sc2 As Double
    sc1 = rst1!nc / 1000000
    sc2 = ntts / 1000000
    
    sc1 = Math.Round(sc1)
    sc2 = Math.Round(sc2)

    sc1 = sc1 * 1000000
    sc2 = sc2 * 1000000
        
        For i = 1 To 100 Step 0.1
            prft = (sc1 * i) / 100
            Dim psc As Double
            psc = 0
            psc = prft / 1000000
            psc = Math.Round(psc)
            psc = psc * 1000000
            If psc = sc2 Then
            Me.Command41.Caption = Math.Round(i, 3) & "%"
            End If
        
        Next i

    End If
    
End Sub

Private Sub Command41_Click()

End Sub

Private Sub Command41_DblClick(Cancel As Integer)
Text23.Visible = True

End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500
Text23.Visible = False
End Sub
Private Sub Text26_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)

End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
KeyAscii = keydwn(KeyAscii)

End Sub

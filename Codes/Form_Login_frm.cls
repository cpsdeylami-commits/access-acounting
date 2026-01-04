VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Login_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command17_Click()
    
    If IsNull(Me.Text12.Value) Or IsNull(Me.Text14.Value) Then
        MsgBox "Please Enter your Information !", vbInformation
    Else
        If Me.Text12.Value = "admin" And Me.Text14.Value = "2024" Then
        DoCmd.Close acForm, "Login_frm"
        DoCmd.OpenForm "popupform"
        Else
        MsgBox "Your ID or Password is Incorrect !", vbInformation
        End If
    End If
End Sub

Private Sub Command18_Click()
DoCmd.Quit acQuitSaveAll
End Sub

Private Sub Form_Close()
End Sub

Private Sub Form_Load()
DoCmd.MoveSize 10200, 3500

End Sub

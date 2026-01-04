VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_srch_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Option21.DefaultValue = 1
txtsrchnam.Enabled = False
DoCmd.MoveSize 10200, 3500
Text30.Enabled = False
Command33.Enabled = False

End Sub

Private Sub Option21_Click()
srchtxt.Enabled = True
txtsrchnam.Enabled = False
Text30.Enabled = False
Command33.Enabled = False

End Sub

Private Sub Option23_Click()
srchtxt.Enabled = False
txtsrchnam.Enabled = True
Text30.Enabled = False
Command33.Enabled = False

End Sub

Private Sub Option31_Click()
srchtxt.Enabled = False
txtsrchnam.Enabled = False
Text30.Enabled = True
Command33.Enabled = True

End Sub

Private Sub srchtxt_KeyPress(KeyAscii As Integer)
'''KeyAscii = keydwn(KeyAscii)
End Sub

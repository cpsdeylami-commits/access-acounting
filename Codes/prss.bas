Attribute VB_Name = "prss"
Option Compare Database

Public Function keydwn(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 8) And Not (KeyAscii = 9) Then
KeyAscii = 0
MsgBox "Please Enter Numbers !", vbInformation
End If
keydwn = KeyAscii
End Function

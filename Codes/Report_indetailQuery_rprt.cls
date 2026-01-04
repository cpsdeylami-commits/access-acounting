VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_indetailQuery_rprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
lbid.Caption = Form_Insrch_frm.inidsrch.Value
Dim t As Long
t = lbid.Caption
If IsNull(DLookup("inid", "indetail", "inid=" & t)) Then
MsgBox "No Invoice Found !", vbInformation
Exit Sub
Else
lbnam.Caption = CStr(DLookup("cnam", "INVOICE", "ID=" & lbid.Caption))
lbph.Caption = CStr(DLookup("CID", "INVOICE", "ID=" & lbid.Caption))
lbadd.Caption = CStr(DLookup("addrs", "cstmr", "phone=" & lbph.Caption))
lbdate.Caption = CStr(DLookup("indate", "INVOICE", "ID=" & lbid.Caption))
End If


''' total invoice calc
Dim d As Integer
Dim s As Double
d = CInt(lbid.Caption)
s = DSum("pt", "indetail", "inid =" & d)
If IsNull(s) Or s = 0 Then
Report_indetailQuery_rprt.inttl.Value = "0"
Else
Report_indetailQuery_rprt.inttl.Value = s
Dim orn As Double
Dim frn As String
orn = inttl.Value
frn = Format(orn, "#,###")
inttl.Value = frn
End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_popupform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command113_Click()
Dim respond As Integer
respond = MsgBox("For Avoid any issue about Update progress we offer you to keep 'Backup' First !", vbOKCancel + vbInformation)
If respond = vbOK Then
UpdateVersion
Else
Exit Sub
End If

End Sub

Private Sub Command31_Click()

    If Me.jari_subform.Visible = False Then
    
        Me.jari_subform.Visible = True
        Box46.Visible = True
        Me.jari_subform_Label.Visible = True
        Me.etc_subform.Visible = False
        Me.etc_subform_Label.Visible = False
    Else
        Me.jari_subform.Visible = False
        Box46.Visible = False
        Me.jari_subform_Label.Visible = True
    
    End If
    
End Sub

Private Sub Command32_Click()

    If Me.etc_subform.Visible = False Then
    
        Me.etc_subform.Visible = True
        Box46.Visible = True
        Me.etc_subform_Label.Visible = True
        Me.jari_subform.Visible = False
        Me.jari_subform_Label.Visible = False
    Else
        Me.etc_subform.Visible = False
        Box46.Visible = False
        Me.etc_subform_Label.Visible = False
    
    End If
    
End Sub

Private Sub Command38_Click()

End Sub

Private Sub Command38_KeyUp(KeyCode As Integer, Shift As Integer)
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command70_Click()
'''Me.Label109.Visible = True

If Me.Command70.Caption = "Show" Then
        
        Dim q As Double
        Dim va As Double
        Dim j As Double
        Dim e As Double
        Dim s As Double
        Dim vs As Double
        
        '''q
        If IsNull(DLookup("Quntity", "Store")) Then
            Me.Label104.Caption = "0"
        Else
            q = DSum("Quntity", "Store")
            Me.Label104.Caption = q
        End If
        
        '''va
        If IsNull(DLookup("ValueOA", "Store")) Then
            Me.Label103.Caption = "0"
        Else
            va = DSum("ValueOA", "Store")
            Me.Label103.Caption = va
        End If
        
        '''j
        If IsNull(DLookup("eql", "jari")) Then
            Me.Label106.Caption = "0"
        Else
            j = DSum("eql", "jari")
            Me.Label106.Caption = j
        End If
        
        '''e
        If IsNull(DLookup("eqletc", "etc")) Then
            Me.Label105.Caption = "0"
        Else
            e = DSum("eqletc", "etc")
            Me.Label105.Caption = e
        End If
        
        '''s
        If IsNull(DLookup("OutQuantity", "Sale")) Then
            Me.Label108.Caption = "0"
        Else
            s = DSum("OutQuantity", "Sale")
            Me.Label108.Caption = s
        End If
        
        '''vs
        If IsNull(DLookup("TotalP", "Sale")) Then
            Me.Label107.Caption = "0"
        Else
            vs = DSum("TotalP", "Sale")
            Me.Label107.Caption = vs
        End If

        
        Me.Box69.Visible = True
        Me.Image63.Visible = True
        Me.Command70.Caption = "Hide"
    Else
        Me.Box69.Visible = False
        Me.Image63.Visible = False
        Me.Command70.Caption = "Show"
End If
        
Dim orn As String
Dim frn As String
orn = Me.Label103.Caption
frn = Format(orn, "#,###")
Me.Label103.Caption = frn

orn = Me.Label107.Caption
frn = Format(orn, "#,###")
Me.Label107.Caption = frn

orn = Me.Label108.Caption
frn = Format(orn, "#,###")
Me.Label108.Caption = frn

orn = Me.Label106.Caption
frn = Format(orn, "#,###")
Me.Label106.Caption = frn

orn = Me.Label105.Caption
frn = Format(orn, "#,###")
Me.Label105.Caption = frn

orn = Me.Label104.Caption
frn = Format(orn, "#,###")
Me.Label104.Caption = frn

orn = Me.Label109.Caption
frn = Format(orn, "#,###")
Me.Label109.Caption = frn

End Sub

Private Sub Command8_Click()
Call RunPythonScript
MsgBox "Backup Created Successfuly !", vbInformation
End Sub

Private Sub etc_subform_Exit(Cancel As Integer)


End Sub

Private Sub etc_subform_Label_DblClick(Cancel As Integer)

End Sub

Private Sub Form_Close()
DoCmd.Quit acQuitSaveAll
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       jari_subform.Visible = False
       etc_subform.Visible = False
    End If
End Sub

Private Sub Form_Load()
DoCmd.Maximize
DoCmd.OpenForm "Login_frm"

If (DLookup("versionupdate", "updtvrsn")) > 0 Then
Me.Command113.Enabled = False
End If

Dim dm As DAO.Database
Dim rst As DAO.Recordset
Dim d As String
Dim m As String
Dim dt As Date
Dim tm As Date
Dim daym As String
Dim monm As String
Dim yrm As Integer
Dim nd As Integer
Dim i As Integer

    dt = Date
                            
            nd = Month(dt)
            If nd >= 9 And nd <= 3 Then
                    i = 21
            Else
                    i = 20
            End If
            
            daym = ((Day(dt) - i) * -1)
            monm = Month(dt) - 3
            yrm = 1400 + (Year(dt) - 2000) - 21
            
            If Len(monm) < 2 Then
                m = "0" & monm
                Me.dtlbl.Caption = yrm & "/" & m & "/" & daym
            Else
                m = monm
                Me.dtlbl.Caption = yrm & "/" & m & "/" & daym
            End If
            
            If Len(daym) < 2 Then
                d = "0" & daym
                Me.dtlbl.Caption = yrm & "/" & monm & "/" & d
            Else
                d = daym
                Me.dtlbl.Caption = yrm & "/" & monm & "/" & d
            End If
            
            Dim t As String
            t = yrm & m & d
     
    
     Set db = CurrentDb
     Set rst = db.OpenRecordset("SELECT SUM(TotalP) AS tv FROM Sale WHERE odate =" & t & "")
     If IsNull(rst!tv) Then
        Me.Label109.Caption = "0"
     Else
        Me.Label109.Caption = rst!tv
     End If
     
End Sub

Private Sub Form_Timer()
Me.tmlbl.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub

Private Sub Image63_Click()
DoCmd.OpenForm "Login_frm"
Me.Box69.Visible = False
Me.Command70.Caption = "Show"
End Sub

Private Sub jari_subform_Label_DblClick(Cancel As Integer)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CommandButton1_Click()
    Dim writefile As Boolean
    Dim writedomo As Boolean
    
    
    writefile = False
    writedomo = False
    
    If CheckBox1 Then
        writefile = True
    End If
    If CheckBox2 Then
        writedomo = True
    End If
    
    success = Analyzer(writefile, writedomo)
End Sub

Private Sub CommandButton2_Click()
    Sheet4.Cells.Clear
End Sub

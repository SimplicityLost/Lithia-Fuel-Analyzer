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

Private Sub CommandButton1_Click()
    Dim writefile As Boolean

    writefile = True
    
    Call Analyzer(writefile)
End Sub


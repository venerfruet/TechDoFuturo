VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ListarTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub confirmar_Click()
Dim nomeTabela As String

    If IsNull(Me.listaTabelas) Then Exit Sub
    
    nomeTabela = Me.listaTabelas.column(0)
    
    Me.Visible = False
    
End Sub

Private Sub fechar_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub listaTabelas_DblClick(Cancel As Integer)
    confirmar_Click
End Sub

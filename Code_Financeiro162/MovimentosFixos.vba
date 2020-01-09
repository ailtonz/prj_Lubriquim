Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
End Sub

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click

    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_Pesquisar.lstCadastro.Requery
    DoCmd.Close

Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub codCategoria_Click()
Dim SQL_Definicoes As String
Dim strCategoria As String

Me.codDefinicao = Null

strCategoria = Me.codCategoria

SQL_Definicoes = "SELECT codDefinicao, Definicao, codCategoria " & _
                 " FROM Definicoes where codCategoria = " & strCategoria & _
                 " ORDER BY Definicao"

Me.codDefinicao.RowSource = SQL_Definicoes

If Me.codCategoria.Column(2) Then
    Me.codFuncionario.Enabled = True
    Me.codFuncionario.Locked = False
Else
    Me.codFuncionario.Enabled = False
    Me.codFuncionario.Locked = True
End If

Me.Categoria = Me.codCategoria.Column(1)

End Sub

Private Sub codDefinicao_Click()
    Me.Definicao = Me.codDefinicao.Column(1)
End Sub

Private Sub codFuncionario_Click()
    Me.Descricao = Me.codFuncionario.Column(1)
    Me.Nome = Me.codFuncionario.Column(1)
End Sub


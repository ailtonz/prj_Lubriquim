Option Compare Database

Private Sub DescricaoDoProduto_Click()

    Me.Quantidade = Me.DescricaoDoProduto.Column(2)
    Me.ValorUnitario = Me.DescricaoDoProduto.Column(3)
    Me.Desconto = Me.DescricaoDoProduto.Column(5)
    Me.ValorTotal = Me.DescricaoDoProduto.Column(4)
    
End Sub

Private Sub Funcionario_Click()

    Me.PercentualComissao = Me.Funcionario.Column(2)
    Me.codFuncionario = Me.Funcionario.Column(1)
    
End Sub

Private Sub PercentualComissao_Exit(Cancel As Integer)

    Me.ValorComissao = (Me.ValorTotal * (Me.PercentualComissao - Me.Desconto)) / 100

End Sub

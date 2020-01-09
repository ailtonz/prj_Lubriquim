Option Compare Database
Option Explicit

Private Sub Desconto_Exit(Cancel As Integer)
    
    Me.ValorTotal = Me.ValorTotal - (Me.ValorTotal * Me.Desconto) / 100
    
End Sub

Private Sub DescricaoDoProduto_Click()
    Me.codClassificacao = Me.DescricaoDoProduto.Column(1)
    Me.codSituacao = Me.DescricaoDoProduto.Column(2)
    Me.codTributacao = Me.DescricaoDoProduto.Column(3)
    Me.PercentualComissao = Me.DescricaoDoProduto.Column(5)
End Sub

Private Sub Embalagem_Click()
    Me.Unidade = Me.Embalagem.Column(2)
End Sub

Private Sub Embalagem_GotFocus()
    Me.Embalagem.Requery
End Sub

Private Sub Embalagem_QTD_Exit(Cancel As Integer)
    Me.PesoBruto = Me.Embalagem_QTD * Me.Embalagem.Column(1)
    Me.PesoLiquido = Me.Embalagem_QTD * Me.Embalagem.Column(3)
End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    
    If Not IsNull(Me.DescricaoDoProduto) Then
        Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
        Me.ValorTotal = Me.ValorTotal - (Me.ValorTotal * Me.Desconto) / 100
    End If
        
End Sub

Private Sub ValorUnitario_Exit(Cancel As Integer)
    
    If Not IsNull(Me.DescricaoDoProduto) Then
        Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
        Me.ValorTotal = Me.ValorTotal - (Me.ValorTotal * Me.Desconto) / 100
    End If
    
End Sub


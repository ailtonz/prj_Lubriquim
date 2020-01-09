Option Compare Database
Option Explicit

Private Sub DescricaoDoProduto_Click()
    
    Me.codClassificacao = Me.DescricaoDoProduto.Column(1)
    Me.codSituacao = Me.DescricaoDoProduto.Column(2)
    Me.codTributacao = Me.DescricaoDoProduto.Column(3)
    Me.PesoLiquido = Me.DescricaoDoProduto.Column(4)
    Me.Embalagem = Me.DescricaoDoProduto.Column(6)
    Me.Unidade = Me.DescricaoDoProduto.Column(7)
    Me.Quantidade = Me.DescricaoDoProduto.Column(8)
    Me.ValorUnitario = Me.DescricaoDoProduto.Column(9)
    Me.ValorTotal = Me.DescricaoDoProduto.Column(10)
    Me.IPI = Me.DescricaoDoProduto.Column(11)
    Me.ValorDoIPI = Me.DescricaoDoProduto.Column(12)
    Me.Lote = Me.DescricaoDoProduto.Column(13)
    Me.LoteData = Me.DescricaoDoProduto.Column(14)
    Me.PesoBruto = Me.DescricaoDoProduto.Column(15)
    Me.codProdCliente = Me.DescricaoDoProduto.Column(16)
    Me.Embalagem_QTD = Me.DescricaoDoProduto.Column(17)
    Me.ICMS = Me.DescricaoDoProduto.Column(18)
    
End Sub

Private Sub Embalagem_Click()
    Me.PesoBruto = Me.PesoLiquido + Me.Embalagem.Column(1)
    Me.Unidade = Me.Embalagem.Column(2)
End Sub

Private Sub Embalagem_GotFocus()
    Me.Embalagem.Requery
End Sub

Private Sub Quantidade_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

Private Sub ValorUnitario_Exit(Cancel As Integer)
    If Not IsNull(Me.DescricaoDoProduto) Then Me.ValorTotal = Me.Quantidade * Me.ValorUnitario
End Sub

Option Compare Database
Option Explicit

Private Sub codCFOP_Click()
    Me.NaturezaDeOperacao = Me.codCFOP.Column(1)
End Sub

Private Sub codCliente_Click()
    Me.Cliente = Me.codCliente.Column(1)
End Sub

Private Sub codEntrega_Click()
    Me.codEntrega.Requery
End Sub

Private Sub codPedidoLubriquim_Click()

    Me.codCliente = Me.codPedidoLubriquim.Column(2)
    Me.codEntrega = Me.codPedidoLubriquim.Column(3)
    Me.codCobranca = Me.codPedidoLubriquim.Column(4)
    Me.codTransportadora = Me.codPedidoLubriquim.Column(5)
    Me.codPedidoCliente = Me.codPedidoLubriquim.Column(6)
    Call codCliente_Click

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    
    If Me.NewRecord Then
       Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
       Me.DataDeEmissao = Format(Now(), "dd/mm/yy")
    End If
    
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

Private Sub Lubriquim_NotasFiscaisItens_Exit(Cancel As Integer)
    Me.ValorTotalDosProdutos = Me.txtSomaProdutos
    Me.ValorTotalDaNota = Me.txtSomaProdutos
    Me.BaseDeCalculoDoICMS = Me.txtSomaProdutos
    Me.PesoBruto = Me.txtTotalDePesos
    Me.PesoLiquido = Me.txtSomaPesoLiquido
    Me.Quantidade = Me.txtEmbalagens
    Me.FaturaNumero = Me.Codigo
    Me.FaturaValor = Me.ValorTotalDaNota
    Me.FaturaOrdem = Me.Codigo
    Me.ValorDoICMS = Me.ValorTotalDaNota * Me.codCliente.Column(2) / 100
End Sub

Private Sub cmdFX1050_Click()
On Error GoTo Err_cmdFX1050_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    'Visualizar Documento
    stDocName = "Lubriquim_NF_FX1050"
    DoCmd.OpenReport stDocName, acPreview, , "codNotaFiscal = " & Me.Codigo

Exit_cmdFX1050_Click:
    Exit Sub

Err_cmdFX1050_Click:
    MsgBox Err.Description
    Resume Exit_cmdFX1050_Click
   
End Sub


Private Sub cmdLQ1070_Click()
On Error GoTo Err_cmdLQ1070_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    'Visualizar Documento
    stDocName = "Lubriquim_NF_LQ1070"
    DoCmd.OpenReport stDocName, acPreview, , "codNotaFiscal = " & Me.Codigo

Exit_cmdLQ1070_Click:
    Exit Sub

Err_cmdLQ1070_Click:
    MsgBox Err.Description
    Resume Exit_cmdLQ1070_Click

End Sub

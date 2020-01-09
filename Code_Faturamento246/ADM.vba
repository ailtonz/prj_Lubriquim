Option Compare Database
Option Explicit

Public Function Comissao(valParcela, valAvista, valComissao)
    Comissao = (valComissao * (valParcela / valAvista * 100)) / 100
End Function

Public Function ExportarDuplicatas(val As String)
On Error GoTo Err_ExportarDuplicatas

Dim Origem As DAO.Recordset

Set Origem = CurrentDb.OpenRecordset("Select * from Lubriquim_Duplicatas where ImportacaoReceita = 0")

If Origem.RecordCount > 0 Then

    Dim dbDestino As DAO.Database
    Dim Destino As DAO.Recordset
    Dim rstTabela As DAO.Recordset
    
    Dim objAccess As Object
    
    Dim strArq As String
    Dim strSQL As String
    Dim codFaturamento As Long
    Dim cont As Long
    Dim s As Variant
    Dim I As Long
    
    Set objAccess = CreateObject("Access.Application")
    
    strArq = AbrirArquivo("Informe o destino das duplicatas", "BDs do Access", "*.MDB;*.MDE", False)
    
    
    'Se selecionou arquivo, atualiza os vínculos
    If strArq <> "" Then
        
        objAccess.OpenCurrentDatabase (strArq)
        objAccess.Visible = False
        
        Set dbDestino = objAccess.CurrentDb
        Set Destino = dbDestino.OpenRecordset("Select * from Faturamentos")
            
        Set rstTabela = dbDestino.OpenRecordset("SELECT Max([codFaturamento])+1 AS CodigoNovo FROM Faturamentos;")
        
        codFaturamento = IIf(IsNull(rstTabela.Fields("CodigoNovo")), 1, rstTabela.Fields("CodigoNovo"))
            
        Origem.MoveLast
        cont = Origem.RecordCount
        Origem.MoveFirst
        
        BeginTrans
        
        s = SysCmd(acSysCmdInitMeter, "Exportando " & cont & " Duplicatas", cont)
        
        I = 0
        
        DoCmd.Hourglass True
        
        While Not Origem.EOF
                            
            Destino.AddNew
            Destino.Fields("codFaturamento") = codFaturamento
            Destino.Fields("codNotaFiscal") = Origem.Fields("codNotaFiscal")
            Destino.Fields("Razao") = Origem.Fields("Apelido")
            Destino.Fields("DataDeEmissao") = Origem.Fields("Emissao")
            Destino.Fields("DataDeVencimento") = Origem.Fields("Vencimento")
            Destino.Fields("ValorDoFaturamento") = Origem.Fields("Valor")
            Destino.Fields("DescricaoDoFaturamento") = Origem.Fields("Ordem")
            Destino.Fields("codTipoFaturamento") = "Faturamento"
            Destino.Fields("Status") = "Aberto"
            Destino.Update
            
            codFaturamento = codFaturamento + 1
            
            Origem.Edit
            Origem.Fields("ImportacaoReceita") = -1
            Origem.Update
            
            Origem.MoveNext
            
            I = I + 1
            
            s = SysCmd(acSysCmdUpdateMeter, I)
            
        Wend
        
        CommitTrans
        
        objAccess.CloseCurrentDatabase
    
        objAccess.Quit
        
        Set objAccess = Nothing
        Set Origem = Nothing
        Set Destino = Nothing
        Set dbDestino = Nothing
        Set rstTabela = Nothing
        
        DoCmd.Hourglass False
        ' Hide the Meter
        s = SysCmd(acSysCmdSetStatus, " ")
        
        MsgBox "Duplicatas enviadas com sucesso.", vbOKOnly + vbInformation, "Exportação de Duplicatas ao financeiro"
    
    End If

Else
    Set Origem = Nothing
    DoCmd.Hourglass False
    MsgBox "Não há Duplicatas a serem enviadas ao financeiro.", vbOKOnly + vbExclamation, "Exportação de Duplicatas ao financeiro"
    Exit Function
End If

Exit_ExportarDuplicatas:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarDuplicatas:
    MsgBox Err.Description
    Resume Exit_ExportarDuplicatas

End Function


Public Function ExportarComissoes()
    On Error GoTo Err_ExportarComissoes

Dim Origem As DAO.Recordset
Dim strOrigem As String

Dim Duplicatas As DAO.Recordset
Dim strDuplicatas As String

strOrigem = "SELECT " & _
            "   qryComissao.codPedidoLubriquim, qryComissao.codNotaFiscal, " & _
            "   qryComissao.Funcionario, Sum(qryComissao.ValorComissao) AS ValComissao, qryComissao.ValorTotalDaNota " & _
            "FROM " & _
            "   (SELECT " & _
            "   Lubriquim_PedidosVendasComissoes.Funcionario, Lubriquim_PedidosVendasComissoes.codPedidoLubriquim, " & _
            "   Lubriquim_NotasFiscais.codNotaFiscal, Lubriquim_PedidosVendasComissoes.ValorComissao, " & _
            "   Lubriquim_NotasFiscais.ValorTotalDaNota, Lubriquim_PedidosVendasComissoes.PG " & _
            "FROM " & _
            "   (Lubriquim_PedidosVendas INNER JOIN Lubriquim_PedidosVendasComissoes ON Lubriquim_PedidosVendas.codPedidoLubriquim = Lubriquim_PedidosVendasComissoes.codPedidoLubriquim) " & _
            "   INNER JOIN Lubriquim_NotasFiscais ON Lubriquim_PedidosVendas.codPedidoLubriquim = Lubriquim_NotasFiscais.codPedidoLubriquim " & _
            "WHERE Lubriquim_PedidosVendasComissoes.PG = No) as qryComissao " & _
            "GROUP BY " & _
            "   qryComissao.codPedidoLubriquim, qryComissao.codNotaFiscal, " & _
            "   qryComissao.Funcionario, qryComissao.ValorTotalDaNota " & _
            "ORDER BY " & _
            "   qryComissao.codNotaFiscal"


Set Origem = CurrentDb.OpenRecordset(strOrigem)

If Origem.RecordCount > 0 Then
    DoCmd.Hourglass True
    While Not Origem.EOF
        strDuplicatas = "SELECT  " & _
                        "   Lubriquim_Duplicatas.codDuplicata, Lubriquim_Duplicatas.Apelido,  " & _
                        "   Lubriquim_Duplicatas.codNotaFiscal, Lubriquim_Duplicatas.Ordem, Lubriquim_Duplicatas.Valor " & _
                        "FROM  " & _
                        "   Lubriquim_Duplicatas " & _
                        "WHERE  " & _
                        "   (((Lubriquim_Duplicatas.codNotaFiscal)=" & Origem.Fields("codNotaFiscal") & ")) " & _
                        "ORDER BY  " & _
                        "   Lubriquim_Duplicatas.codDuplicata "
        Set Duplicatas = CurrentDb.OpenRecordset(strDuplicatas)
        If Duplicatas.RecordCount > 0 Then
            While Not Duplicatas.EOF
                MsgBox Duplicatas.Fields("codNotaFiscal")
                MsgBox Comissao(Duplicatas.Fields("Valor"), Origem.Fields("ValorTotalDaNota"), Origem.Fields("ValComissao"))
                Duplicatas.MoveNext
            Wend
        End If
        Origem.MoveNext
    Wend

End If

Set Origem = Nothing
Set Duplicatas = Nothing


Exit_ExportarComissoes:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarComissoes:
    MsgBox Err.Description
    Resume Exit_ExportarComissoes
End Function

Public Function ExportarMateriaPrima()
On Error GoTo Err_ExportarMateriaPrima

Dim Origem As DAO.Recordset

Set Origem = CurrentDb.OpenRecordset("Select * from FAT_ProdutosMateriasPrimas")

If Origem.RecordCount > 0 Then

    Dim dbDestino As DAO.Database
    Dim Destino As DAO.Recordset
    Dim rstTabela As DAO.Recordset
    
    Dim objAccess As Object
    
    Dim strArq As String
    Dim strSQL As String
    Dim codFaturamento As Long
    Dim cont As Long
    Dim s As Variant
    Dim I As Long
    
    Set objAccess = CreateObject("Access.Application")
    
    strArq = AbrirArquivo("Informe o destino", "BDs do Access", "*.MDB;*.MDE", False)
    
    
    'Se selecionou arquivo, atualiza os vínculos
    If strArq <> "" Then
        
        objAccess.OpenCurrentDatabase (strArq)
        objAccess.Visible = False
        
        Set dbDestino = objAccess.CurrentDb
        Set Destino = dbDestino.OpenRecordset("Select * from FAT_Produtos")
            
        Set rstTabela = dbDestino.OpenRecordset("SELECT Max([codProduto])+1 AS CodigoNovo FROM FAT_Produtos;")
        
        codFaturamento = IIf(IsNull(rstTabela.Fields("CodigoNovo")), 1, rstTabela.Fields("CodigoNovo"))
            
        Origem.MoveLast
        cont = Origem.RecordCount
        Origem.MoveFirst
        
        BeginTrans
        
        s = SysCmd(acSysCmdInitMeter, "Exportando " & cont, cont)
        
        I = 0
        
        DoCmd.Hourglass True
        
        While Not Origem.EOF
                            
            Destino.AddNew
            Destino.Fields("codProduto") = codFaturamento
            Destino.Fields("DescricaoDoProduto") = Origem.Fields("MateriaPrima")
            Destino.Fields("codInterno") = Origem.Fields("CodigoInterno")
            Destino.Fields("codTipo") = "MP"
            Destino.Fields("Valor") = Origem.Fields("Valor")
            Destino.Update
            
            codFaturamento = codFaturamento + 1
                        
            Origem.MoveNext
            
            I = I + 1
            
            s = SysCmd(acSysCmdUpdateMeter, I)
            
        Wend
        
        CommitTrans
        
        objAccess.CloseCurrentDatabase
    
        objAccess.Quit
        
        Set objAccess = Nothing
        Set Origem = Nothing
        Set Destino = Nothing
        Set dbDestino = Nothing
        Set rstTabela = Nothing
        
        DoCmd.Hourglass False
        ' Hide the Meter
        s = SysCmd(acSysCmdSetStatus, " ")
        
        MsgBox "Enviadas com sucesso.", vbOKOnly + vbInformation, "Exportação"
    
    End If

Else
    Set Origem = Nothing
    DoCmd.Hourglass False
    MsgBox "Não há itens a serem enviados.", vbOKOnly + vbExclamation, "Exportação"
    Exit Function
End If

Exit_ExportarMateriaPrima:
    DoCmd.Hourglass False
    Exit Function

Err_ExportarMateriaPrima:
    MsgBox Err.Description
    Resume Exit_ExportarMateriaPrima

End Function

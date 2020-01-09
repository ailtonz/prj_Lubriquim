Option Compare Database
Option Explicit

Public Inicio As String
Public Final As String

Public Function ImportarMovimentosFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Movimentos As DAO.Recordset
Dim ItensDeMovimentos As DAO.Recordset

Dim MovimentosFixos As DAO.Recordset
Dim ItensDeMovimentosFixos As DAO.Recordset


Set Movimentos = CurrentDb.OpenRecordset("Select * from Movimentos")
Set ItensDeMovimentos = CurrentDb.OpenRecordset("Select * from ItensDeMovimentos")
Set MovimentosFixos = CurrentDb.OpenRecordset("Select * from MovimentosFixos where MovimentoAtivo = true")



BeginTrans

While Not MovimentosFixos.EOF
    Movimentos.AddNew
    
    ' DADOS GERAIS
    Movimentos.Fields("codMovimento") = NovoCodigo("Movimentos", "codMovimento")
    Movimentos.Fields("DataDeEmissao") = Format(Now(), "dd/mm/yy")
    Movimentos.Fields("DescricaoDoMovimento") = MovimentosFixos.Fields("Descricao")
    Movimentos.Fields("codCategoria") = MovimentosFixos.Fields("codCategoria")
    Movimentos.Fields("codDefinicao") = MovimentosFixos.Fields("codDefinicao")
    Movimentos.Fields("codFuncionario") = MovimentosFixos.Fields("codFuncionario")
    Movimentos.Fields("ValorDoMovimento") = MovimentosFixos.Fields("ValorDoMovimento")
    Movimentos.Fields("DataDeVencimento") = CalcularVencimento(MovimentosFixos.Fields("DiaDeVencimento"), MES, ANO)
    
    ' STATUS = "Aberto"
    Movimentos.Fields("codStatus") = 2
    Movimentos.Fields("Status") = "Aberto"
    
    ' TIPO DE MOVIMENTO
    Movimentos.Fields("codTipoMovimento") = MovimentosFixos.Fields("codTipoMovimento")
    
    ' EXTRAS
    Movimentos.Fields("Categoria") = MovimentosFixos.Fields("Categoria")
    Movimentos.Fields("Definicao") = MovimentosFixos.Fields("Definicao")
    Movimentos.Fields("Nome") = MovimentosFixos.Fields("Nome")
    
    Movimentos.Update
    Movimentos.MoveLast
    
    Set ItensDeMovimentosFixos = CurrentDb.OpenRecordset _
        ("Select * from ItensDeMovimentosFixos where codMovimentoFixo = " & _
        MovimentosFixos.Fields("codMovimentoFixo"))

    While Not ItensDeMovimentosFixos.EOF
        ItensDeMovimentos.AddNew
        ItensDeMovimentos.Fields("codMovimento") = Movimentos.Fields("codMovimento")
        ItensDeMovimentos.Fields("codDepartamento") = ItensDeMovimentosFixos.Fields("codDepartamento")
        ItensDeMovimentos.Fields("Valor") = ItensDeMovimentosFixos.Fields("Valor")
        ItensDeMovimentos.Fields("TipoDeValor") = ItensDeMovimentosFixos.Fields("TipoDeValor")
        ItensDeMovimentos.Update
        ItensDeMovimentosFixos.MoveNext
    Wend
    MovimentosFixos.MoveNext
Wend

CommitTrans

MsgBox "Importação concluída.", vbInformation, "Importar Movimentos Fixos"

Movimentos.Close
ItensDeMovimentos.Close

MovimentosFixos.Close
'ItensDeMovimentosFixos.Close

End Function


Public Function ImportarFaturamentosFixos(Optional MES As Integer, Optional ANO As Integer)

Dim Faturamentos As DAO.Recordset
Dim ItensDeFaturamentos As DAO.Recordset

Dim FaturamentosFixos As DAO.Recordset
Dim ItensDeFaturamentosFixos As DAO.Recordset


Set Faturamentos = CurrentDb.OpenRecordset("Select * from Faturamentos")
Set ItensDeFaturamentos = CurrentDb.OpenRecordset("Select * from ItensDeFaturamentos")

Set FaturamentosFixos = CurrentDb.OpenRecordset("Select * from FaturamentosFixos Where FaturamentoAtivo = True")

BeginTrans

While Not FaturamentosFixos.EOF
    Faturamentos.AddNew
    
    ' DADOS GERAIS
    Faturamentos.Fields("codFaturamento") = NovoCodigo("Faturamentos", "codFaturamento")
    Faturamentos.Fields("codCliente") = FaturamentosFixos.Fields("codCliente")
    Faturamentos.Fields("DescricaoDoFaturamento") = FaturamentosFixos.Fields("Descricao")
    Faturamentos.Fields("ValorDoFaturamento") = FaturamentosFixos.Fields("ValorDoFaturamento")
    Faturamentos.Fields("DataDeVencimento") = CalcularVencimento(FaturamentosFixos.Fields("DiaDeVencimento"), MES, ANO)
    
    ' STATUS = "Aberto"
    Faturamentos.Fields("codStatus") = 2
    
    ' TIPO DE FATURAMENTO
    Faturamentos.Fields("codTipoFaturamento") = FaturamentosFixos.Fields("codTipoFaturamento")
    
    ' EXTRAS
    Faturamentos.Fields("Razao") = FaturamentosFixos.Fields("Razao")
    
    Faturamentos.Update
    Faturamentos.MoveLast
    
    Set ItensDeFaturamentosFixos = CurrentDb.OpenRecordset _
        ("Select * from ItensDeFaturamentosFixos where codFaturamentoFixo = " & _
        FaturamentosFixos.Fields("codFaturamentoFixo"))

    While Not ItensDeFaturamentosFixos.EOF
        ItensDeFaturamentos.AddNew
        ItensDeFaturamentos.Fields("codFaturamento") = Faturamentos.Fields("codFaturamento")
        ItensDeFaturamentos.Fields("codDepartamento") = ItensDeFaturamentosFixos.Fields("codDepartamento")
        ItensDeFaturamentos.Fields("Valor") = ItensDeFaturamentosFixos.Fields("Valor")
        ItensDeFaturamentos.Fields("TipoDeValor") = ItensDeFaturamentosFixos.Fields("TipoDeValor")
        ItensDeFaturamentos.Update
        ItensDeFaturamentosFixos.MoveNext
    Wend
    FaturamentosFixos.MoveNext
Wend

CommitTrans

MsgBox "Importação concluída.", vbInformation, "Importar Faturamentos Fixos"

Faturamentos.Close
ItensDeFaturamentos.Close

FaturamentosFixos.Close
'ItensDeFaturamentosFixos.Close

End Function


Public Function CalcularVencimento(DIA As Integer, Optional MES As Integer, Optional ANO As Integer) As Date

If MES > 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, MES, DIA)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO > 0 Then
    CalcularVencimento = Format((DateSerial(ANO, Month(Now), DIA)), "dd/mm/yyyy")
ElseIf MES = 0 And ANO = 0 Then
    CalcularVencimento = Format((DateSerial(Year(Now), Month(Now), DIA)), "dd/mm/yyyy")
End If

End Function

Public Function Zebrar(rpt As Report)
Static fCinza As Boolean
Const conCinza = 15198183
Const conBranco = 16777215

On Error Resume Next

    rpt.Section(0).BackColor = IIf(fCinza, conCinza, conBranco)
    fCinza = Not fCinza

End Function


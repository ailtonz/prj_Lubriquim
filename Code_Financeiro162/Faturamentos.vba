Option Compare Database
Option Explicit

Private Sub cmdCalcular_Click()
Dim Valor As Currency
    Valor = Me.ValorDoFaturamento
    Me.D_Dias = DateDiff("d", Me.DataDeEmissao, Me.DataDeVencimento)
    Me.D_VLEncargos = Valor / 100 * ((Me.D_Encargos / 30) * Me.D_Dias)
    Me.ValorDoFaturamento = Valor - Me.D_VLEncargos - Me.D_Banco - Me.D_Adicional
    Me.D_Total = Valor
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

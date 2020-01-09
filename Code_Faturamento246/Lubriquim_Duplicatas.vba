Option Compare Database
Option Explicit

Private Sub codCliente_Click()
    Me.Cliente = Me.codCliente.Column(1)
    Me.Apelido = Me.codCliente.Column(2)
End Sub

Private Sub codCobranca_GotFocus()
    Me.codCobranca.Requery
End Sub

Private Sub codNotaFiscal_Click()

    Me.Fatura = Me.codNotaFiscal.Column(1)
    Me.Valor = Me.codNotaFiscal.Column(2)
    Me.Ordem = Me.codNotaFiscal.Column(3)
    Me.Vencimento = Me.codNotaFiscal.Column(4)
    Me.codCliente = Me.codNotaFiscal.Column(5)
    Me.codCobranca = Me.codNotaFiscal.Column(6)
    Call codCliente_Click

End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    If Me.NewRecord Then
       Me.Codigo = NovoCodigo(Me.RecordSource, Me.Codigo.ControlSource)
       Me.Emissao = Format(Now(), "dd/mm/yy")
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

Private Sub cmdFX1050_Click()
On Error GoTo Err_cmdFX1050_Click

    Dim stDocName As String
    
    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    stDocName = "Lubriquim_DPL_FX1050"
    DoCmd.OpenReport stDocName, acPreview, , "codDuplicata = " & Me.Codigo

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
    
    stDocName = "Lubriquim_DPL_LQ1070"
    DoCmd.OpenReport stDocName, acPreview, , "codDuplicata = " & Me.Codigo

Exit_cmdLQ1070_Click:
    Exit Sub

Err_cmdLQ1070_Click:
    MsgBox Err.Description
    Resume Exit_cmdLQ1070_Click

End Sub

Option Compare Database
Option Explicit

Private Sub codCliente_Click()
    Me.UGE = Me.codCliente.Column(2)
End Sub

Private Sub codEndereco_GotFocus()
End Sub

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

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click

    Dim stDocName As String
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    
    stDocName = "ArtTecnica_PedidosVendas"
    DoCmd.OpenReport stDocName, acPreview, , "ArtTecnica_PedidosVendas.codPedidoArtTecnica = " & Me.Codigo

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub



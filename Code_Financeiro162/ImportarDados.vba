Option Compare Database
Option Explicit

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdDesfazer_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close
    
Exit_cmdDesfazer_Click:
    Exit Sub

Err_cmdDesfazer_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdDesfazer_Click

End Sub

Private Sub cmdFatFixos_Click()

Select Case optCriterio.Value

    Case 1
    
        ImportarFaturamentosFixos
    
    Case 2
        
        ImportarFaturamentosFixos Me.cboMes.Column(0), Me.txtAno.Value
    
End Select


End Sub


Private Sub cmdMovFixos_Click()
    
Select Case optCriterio.Value

    Case 1
    
        ImportarMovimentosFixos
    
    Case 2
        
        ImportarMovimentosFixos Me.cboMes.Column(0), Me.txtAno.Value
    
End Select

End Sub

Private Sub optCriterio_Click()

Select Case optCriterio.Value

    Case 1
    
        Me.cboMes.Enabled = False
        Me.txtAno.Enabled = False
    
    Case 2
        
        Me.cboMes.Enabled = True
        Me.txtAno.Enabled = True
    
End Select


End Sub

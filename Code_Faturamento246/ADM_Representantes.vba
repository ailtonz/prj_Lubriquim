Option Compare Database

Private Sub cboCliente_Click()
Dim strSQL As String

If Me.cboCliente <> "" Then
    strSQL = "SELECT ADM_ClientesXFuncionarios.codRelacao, ADM_ClientesXFuncionarios.codCliente, ADM_ClientesXFuncionarios.codFuncionario FROM ADM_ClientesXFuncionarios WHERE (((ADM_ClientesXFuncionarios.codCliente)=[forms].[ADM_Representantes].[cboCliente]))"
    Me.subRepresentantes.Form.RecordSource = strSQL
Else
    strSQL = "SELECT ADM_ClientesXFuncionarios.codRelacao, ADM_ClientesXFuncionarios.codCliente, ADM_ClientesXFuncionarios.codFuncionario FROM ADM_ClientesXFuncionarios"
    Me.subRepresentantes.Form.RecordSource = strSQL
End If

Me.cboRepresentante.value = Null
Me.subRepresentantes.Requery

End Sub

Private Sub cboRepresentante_Click()
Dim strSQL As String

If Me.cboRepresentante <> "" Then
    strSQL = "SELECT ADM_ClientesXFuncionarios.codRelacao, ADM_ClientesXFuncionarios.codCliente, ADM_ClientesXFuncionarios.codFuncionario FROM ADM_ClientesXFuncionarios WHERE (((ADM_ClientesXFuncionarios.codFuncionario)=[forms].[ADM_Representantes].[cboRepresentante]))"
    Me.subRepresentantes.Form.RecordSource = strSQL
Else
    strSQL = "SELECT ADM_ClientesXFuncionarios.codRelacao, ADM_ClientesXFuncionarios.codCliente, ADM_ClientesXFuncionarios.codFuncionario FROM ADM_ClientesXFuncionarios"
    Me.subRepresentantes.Form.RecordSource = strSQL
End If
Me.cboCliente.value = Null
Me.subRepresentantes.Requery
End Sub

'Private Sub cmdTroca_Click()
'Dim strSQL As String
'
'strSQL = "update ADM_ClientesXFuncionarios set codFuncionario = " & Me.cboPara.Column(0) & " where codFuncionario = " & Me.cboDE.Column(0) & ""
'
'DoCmd.RunSQL strSQL
'
'Me.subRepresentantes.Requery
'Me.cboRepresentante.Requery
'Me.cboDE.Requery
'
'End Sub

Private Sub Form_Load()

    Me.subRepresentantes.Requery

End Sub

Private Sub Form_Resize()
Dim X

X = RedimencionaControle(Me, [subRepresentantes])

End Sub

Private Sub btnCrecente_Click()
On Error GoTo Error_btnCrecente

    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdSortAscending
    
Exit_btnCrecente:
    Exit Sub

Error_btnCrecente:
    MsgBox Err & ": " & Err.Description
    Resume Exit_btnCrecente

End Sub

Private Sub btnDecrecente_Click()
On Error GoTo Error_btnDecrecente

    Screen.PreviousControl.SetFocus
    DoCmd.RunCommand acCmdSortDescending

Exit_btnDecrecente:
    Exit Sub

Error_btnDecrecente:
    MsgBox Err & ": " & Err.Description
    Resume Exit_btnDecrecente
End Sub

Private Sub subRepresentantes_Exit(Cancel As Integer)

'Me.cboDE.Requery
'Me.cboPara.Requery
Me.cboRepresentante.Requery
Me.cboCliente.Requery

End Sub

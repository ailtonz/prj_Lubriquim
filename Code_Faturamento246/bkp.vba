Option Compare Database

Private Sub cmdBKP_Click()
    bkp
End Sub

Private Sub cmdCaminho_Click()
Dim strCaminho As String
strCaminho = CaminhoDoBKP

If Len(strCaminho) > 0 Then
    If Right(strCaminho, 1) = "\" Then
        Me.txtCaminho = strCaminho
    Else
        Me.txtCaminho = strCaminho & "\"
    End If

    'Salvar Registro
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
End If

End Sub

Private Function CaminhoDoBKP() As String

Dim lngCount As Long

' Open the file dialog
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = True
    '.Filters.Add FiltroTitulo, FiltroExtencao
    .Title = "Indique o caminho para bkp"
    '.AllowMultiSelect = MultiSelecao
    .Show
    
    ' Display paths of each file selected
    For lngCount = 1 To .SelectedItems.Count
        CaminhoDoBKP = .SelectedItems(lngCount)
    Next lngCount
    
End With

End Function


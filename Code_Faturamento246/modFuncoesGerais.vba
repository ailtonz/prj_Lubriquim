Option Compare Database
Option Explicit

Public strTabela As String

Public Function LocalizarBanco() As String
Dim Arq As String
Dim Caminho As String
Dim Banco As String


Arq = "caminho.log"
Caminho = Application.CurrentProject.Path & "\" & Arq

If VerificaExistenciaDeArquivo(Caminho) Then
    Banco = getCaminho(Application.CurrentProject.Path & "\" & Arq)
    If VerificaExistenciaDeArquivo(Banco) Then
        LocalizarBanco = Banco
    Else
        MsgBox "ATENÇÃO: Não é possível localizar o Banco de dados.", vbExclamation + vbOKOnly
        LocalizarBanco = ""
        'Apaga Caminho invalido!
        Kill Caminho
    End If
Else
    MsgBox "ATENÇÃO: Não é possível localizar o Banco de dados.", vbExclamation + vbOKOnly
    LocalizarBanco = ""
End If

End Function

Private Function VerificaExistenciaDeArquivo(Localizacao As String) As Boolean

If Dir(Localizacao, vbArchive) <> "" Then
    VerificaExistenciaDeArquivo = True
Else
    VerificaExistenciaDeArquivo = False
End If

End Function

Private Function getCaminho(arqCaminho As String) As String
Dim lin As String

Open arqCaminho For Input As #1

Line Input #1, lin
getCaminho = lin

Close #1

End Function


Public Function ExecutarSQL(strSQL As String)

'Desabilitar menssagens de execução de comando do access
DoCmd.SetWarnings False

DoCmd.RunSQL strSQL

'Abilitar menssagens de execução de comando do access
DoCmd.SetWarnings True

End Function

Public Function GerarSaida(strConteudo As String, strArquivo As String)

Open Application.CurrentProject.Path & "\" & strArquivo For Append As #1

Print #1, strConteudo

Close #1

End Function

Public Function CriarPasta(sPasta As String) As String
'Cria pasta apartir da origem do sistema

Dim fPasta As New FileSystemObject
Dim MyApl As String

MyApl = Application.CurrentProject.Path

If Not fPasta.FolderExists(MyApl & "\" & sPasta) Then
   fPasta.CreateFolder (MyApl & "\" & sPasta)
End If

CriarPasta = MyApl & "\" & sPasta & "\"

End Function

Public Function getPath(sPathIn As String) As String
'Esta função irá retornar apenas o path de uma string que contenha o path e o nome do arquivo:
Dim I As Integer

  For I = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, I, 1)) Then Exit For
  Next

  getPath = Left$(sPathIn, I)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next

  getFileName = Left(Mid$(sFileIn, I + 1, Len(sFileIn) - I), Len(Mid$(sFileIn, I + 1, Len(sFileIn) - I)) - 4)

End Function

Public Function getFileExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next

  getFileExt = Right(Mid$(sFileIn, I + 1, Len(sFileIn) - I), 4)

End Function

Public Function RedimencionaControle(frm As Form, ctl As Control)

Dim intAjuste As Integer
On Error Resume Next

intAjuste = frm.Section(acHeader).Height * frm.Section(acHeader).Visible

On Error GoTo 0

intAjuste = Abs(intAjuste) + ctl.Top

If intAjuste < frm.InsideHeight Then
    ctl.Height = frm.InsideHeight - intAjuste
End If

End Function

Public Function NovoCodigo(Tabela, Campo)

Dim rstTabela As DAO.Recordset
Set rstTabela = CurrentDb.OpenRecordset("SELECT Max([" & Campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rstTabela.EOF Then
   NovoCodigo = rstTabela.Fields("CodigoNovo")
   If IsNull(NovoCodigo) Then
      NovoCodigo = 1
   End If
Else
   NovoCodigo = 1
End If
rstTabela.Close

End Function


Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function AbrirArquivo(sTitulo As String, sDecricao As String, sTipo As String, SelecaoMultipla As Boolean) As String
Dim fd As Office.FileDialog

'Diálogo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'Título
fd.Title = sTitulo

'Filtros e descrição dos mesmos
fd.Filters.Add sDecricao, sTipo

'Premissões de selação
fd.AllowMultiSelect = SelecaoMultipla

If fd.Show = -1 Then
    AbrirArquivo = fd.SelectedItems(1)
End If

End Function

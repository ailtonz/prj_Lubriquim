Option Compare Database
Public Function bkp()
Dim rstBKP As DAO.Recordset

Set rstBKP = CurrentDb.OpenRecordset("Select * from bkp order by codCaminho")

While Not rstBKP.EOF
    Backup rstBKP.Fields("Caminho")
    rstBKP.MoveNext
Wend

rstBKP.Close

End Function


Public Function Backup(sDestino As String)
'===================================================================
'   Funções agregadas a esta função:
'   > CompactarRepararDatabase
'   > CriarPasta
'   > getPath
'   > getFileName
'   > getFileExt
''===================================================================

Dim oFSO As New FileSystemObject
Dim oPasta As New FileSystemObject
Dim oSHL
Dim tmp, p1, p2, p3, p4, p5
Dim Origem As String
Dim sOrigem As String
Dim sArquivo As String
Dim sExtencao As String

sOrigem = Application.CurrentProject.Path
sArquivo = getFileName(Application.CurrentProject.Path & "\db" & Application.CurrentProject.Name)
sExtencao = getFileExt(Application.CurrentProject.Path & "\db" & Application.CurrentProject.Name)

On Error Resume Next
Err.Clear

Origem = sOrigem & "\" & sArquivo & sExtencao

'Começa o bkp se o arquivo existir na origem
If Dir(Origem) <> "" Then
   
    Application.Screen.MousePointer = 11
   
    p1 = Right("00" & Year(Now()), 2)
    p2 = Right("00" & Month(Now()), 2)
    p3 = Right("00" & Day(Now()), 2)
    p4 = Right("00" & Hour(Now()), 2)
    p5 = Right("00" & Minute(Now()), 2)
     
    tmp = ("_" & p1 & p2 & p3 & "_" & p4 & p5)
    
    CompactarRepararDatabase sOrigem & "\" & sArquivo & sExtencao
    
    sOrigem = sOrigem & "\"
    
    oFSO.CopyFile sOrigem & sArquivo & sExtencao, sDestino & sArquivo & tmp & sExtencao, True
    
    Application.Screen.MousePointer = 0
    
Else
    
    MsgBox "ATENÇÃO: Execute esta operação apartir do computador que contém os dados do sistema", vbInformation + vbOKOnly, "Backup"
    
End If

End Function

Public Function CompactarRepararDatabase(DatabasePath As String, Optional Password As String, Optional TempFile As String = "c:\tmp.mdb")
'===================================================================
' Se a versao DAO for anterior a 3.6 , entao devemos usar o método RepairDatabase
' Se a versao DAO for a 3.6 ou superior basta usar a função CompactDatabase
'===================================================================

If DBEngine.Version < "3.6" Then DBEngine.RepairDatabase DatabasePath

'se nao informou um arquivo temporario usa "c:\tmp.mdb"
If TempFile = "" Then TempFile = "c:\tmp.mdb"

'apaga o arquivo temp se existir
If Dir(TempFile) <> "" Then Kill TempFile

'formata a senha no formato ";pwd=PASSWORD" se a mesma existir
If Password <> "" Then Password = ";pwd=" & Password

'compacta a base criando um novo banco de dados
DBEngine.CompactDatabase DatabasePath, TempFile, , , Password

'apaga o primeiro banco de dados
Kill DatabasePath

'move a base compactada para a origem
FileCopy TempFile, DatabasePath

'apaga o arquivo temporario
Kill TempFile

End Function

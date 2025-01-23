
Dim xlApp 
Dim xlWorkbook
Dim caminhoArquivo

' Caminho para o arquivo Excel
caminhoArquivo = "C:\AUTOMACAO\RELATORIOS PARCELAS\RELATORIOS\7001168832.xlsm"

On Error Resume Next ' Inicia o tratamento de erros

' Cria uma instância do Excel
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False  ' O Excel não será visível

' Abre o arquivo Excel
Set xlWorkbook = xlApp.Workbooks.Open(caminhoArquivo)

' Verifica se o arquivo foi aberto com sucesso
If Err.Number <> 0 Then
    WScript.Echo "Erro ao abrir o arquivo: " & Err.Description
    WScript.Quit
End If

' Executa a macro chamada "teste"
xlApp.Run "teste"

' Salva e fecha o arquivo
xlWorkbook.Save
xlWorkbook.Close

' Libera os objetos
Set xlWorkbook = Nothing
xlApp.Quit
Set xlApp = Nothing

Sub ExportarResumo()
    Dim caminhoBase As String
    Dim caminhoMes As String
    Dim mesAtual As String
    Dim dataHoraAtual As String
    Dim nomeArquivo As String
    Dim ordens As String
    Dim cel As Range
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Definir o diretório base
    caminhoBase = "C:\Users\william.goncalves\Desktop\FATURAMENTO\FATURAMENTO - SERVIÇOS\1 - LAFAETE\2024\"
    
    ' Obter o mês atual
    mesAtual = Format(Date, "mm.yyyy")
    
    ' Construir o caminho para o diretório do mês atual
    caminhoMes = caminhoBase & mesAtual & "\CONFERÊNCIA - LISTAGEM DE SERVIÇOS\"
    
    ' Verificar se o diretório do mês atual existe, caso contrário, criar
    If Dir(caminhoMes, vbDirectory) = "" Then
        MkDir caminhoMes
    End If
    
    ' Selecionar a aba "ORDEM DE VENDA"
    Sheets("RESUMO").Select
    
    ' Remover o ponto e vírgula inicial
    ordens = Mid(ordens, 4)
    
    ' Exportar apenas a aba selecionada para o diretório do mês atual
    ActiveSheet.Copy
    dataHoraAtual = Format(Now(), "dd.mm.yyyy hh.mm.ss")
    nomeArquivo = "RESUMO (" & mesAtual & ") " & ordens & ".xlsx"
    ActiveWorkbook.SaveAs caminhoMes & nomeArquivo
    ActiveWorkbook.Close SaveChanges:=False
    
    ' Exibir mensagem de confirmação
    MsgBox "Notas exportadas para: " & caminhoMes & nomeArquivo
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub



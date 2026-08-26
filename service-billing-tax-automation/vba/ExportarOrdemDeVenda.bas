Sub ExportarOrdemDeVenda()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim filePath As String
    Dim fileName As String
    Dim currentDateTime As String
    Dim cell As Range
    
    ' Definindo a aba onde estão as ordens de venda
    Set ws = ThisWorkbook.Sheets("ORDEM DE VENDA")
    
    ' Verificando se a aba existe
    If ws Is Nothing Then
        MsgBox "A aba 'ORDEM DE VENDA' não foi encontrada."
        Exit Sub
    End If
    
    ' Definindo o diretório de destino como a área de trabalho
    filePath = Environ("USERPROFILE") & "\Desktop\FATURAMENTO\FATURAMENTO - SERVIÇOS\1 - LAFAETE\2024\11.2024\"
    
    ' Obtendo a data e hora atual no formato dd.mm.yyyy hh:mm
    currentDateTime = Format(Now, "dd.mm.yyyy hh.mm")
    
    ' Construindo o nome do arquivo
    fileName = "Ordens de Venda (" & currentDateTime & ").xlsx"
    
    ' Definindo o intervalo que será exportado (todas as células utilizadas na aba ORDEM DE VENDA)
    Set rng = ws.UsedRange
    
    ' Exportando o intervalo para um novo arquivo Excel
    rng.Copy
    Workbooks.Add(1).Sheets(1).Paste
    Application.CutCopyMode = False
    
    ' Salvando o arquivo
    Dim newFilePath As String
    newFilePath = filePath & fileName
    ActiveWorkbook.SaveAs newFilePath
    
    ' Fechar o livro
    ActiveWorkbook.Close
    
    MsgBox "A ordem de venda foi exportada com sucesso para a área de trabalho: " & fileName
    
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro ao exportar a ordem de venda."
End Sub



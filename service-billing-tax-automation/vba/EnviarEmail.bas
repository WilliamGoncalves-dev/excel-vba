Sub EnviarEmail()
    ' Variáveis para o Outlook
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    ' Variáveis para extração de e-mails
    Dim ws As Worksheet
    Dim cliente As String
    Dim email As String
    Dim emailsCliente As String ' Variável para armazenar todos os e-mails do cliente
    Dim emailDict As Object ' Dicionário para armazenar e-mails únicos
    
    ' Inicializar dicionário
    Set emailDict = CreateObject("Scripting.Dictionary")
    
    ' Abrir a planilha com os e-mails
    Set ws = ThisWorkbook.Sheets("E-MAIL")
    
    ' Iniciar Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Obter o valor da célula O5 na aba "Tributação fora do Município"
    Dim clienteCelula As Range
    Set clienteCelula = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("O5")
    
    ' Verificar se a célula não está vazia
    If clienteCelula.Value = "" Then
        MsgBox "O valor da célula O5 está vazio. Por favor, insira o valor do cliente."
        Exit Sub
    End If
    
    ' Definir o cliente
    cliente = clienteCelula.Value
    
    ' Extrair e-mails do cliente
    Dim lastRow As Long
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow ' Supondo que os dados começam na linha 2
        If ws.Cells(i, "A").Value = cliente Then
            email = ws.Cells(i, "D").Value
            If email <> "" Then
                If Not emailDict.Exists(email) Then
                    emailDict(email) = True ' Adiciona o e-mail ao dicionário
                    emailsCliente = emailsCliente & email & "; "
                End If
            End If
        End If
    Next i
    
    ' Verificar se encontrou algum e-mail do cliente
    If emailsCliente = "" Then
        MsgBox "E-mails do cliente não encontrados."
        Exit Sub
    End If
    
    ' Obtendo valores das células para concatenar ao assunto do e-mail
    Dim assuntoExtra As String
    assuntoExtra = " OV " & ThisWorkbook.Sheets("Tributação fora do Municipio").Range("F5").Value & " CT " & ThisWorkbook.Sheets("Tributação fora do Municipio").Range("I5").Value & " - " & ThisWorkbook.Sheets("Tributação fora do Municipio").Range("K3").Value
    
    ' Criar e-mail
    Set OutlookMail = OutlookApp.CreateItem(0)
    With OutlookMail
        .To = emailsCliente
        .Subject = "FATURAMENTO - LAFAETE " & assuntoExtra ' Concatenando informações ao assunto do e-mail
        .Display ' Exibir e-mail para revisão
        '.Send ' Enviar e-mail automaticamente
    End With
    
    ' Limpar objetos
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    Set ws = Nothing
    
    ' Copiar informações para a aba RESUMO
    Dim wsResumo As Worksheet
    Set wsResumo = ThisWorkbook.Sheets("RESUMO")
    
    ' Encontrar a próxima linha vazia na aba RESUMO
    Dim nextRow As Long
    nextRow = wsResumo.Cells(wsResumo.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Copiar os valores para a aba RESUMO
    wsResumo.Cells(nextRow, "A").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("F5").Value ' Copiar F5 para A2
    wsResumo.Cells(nextRow, "B").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("I5").Value ' Copiar I5 para B2
    wsResumo.Cells(nextRow, "C").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("K3").Value ' Copiar K3 para C2
    wsResumo.Cells(nextRow, "D").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("G5").Value ' Copiar G5 para D2
    wsResumo.Cells(nextRow, "F").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("D6").Value ' Copiar D4 para F2
    wsResumo.Cells(nextRow, "G").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("D14").Value ' Copiar D14 para G2
    wsResumo.Cells(nextRow, "H").Value = ThisWorkbook.Sheets("Tributação fora do Municipio").Range("K5").Value ' Copiar K5 para H2
    
    ' Inserir a data de processamento na coluna I
    wsResumo.Cells(nextRow, "I").Value = Format(Date, "DD/MM/YYYY")
    
End Sub

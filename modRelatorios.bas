Attribute VB_Name = "modRelatorios"
' ========================================
' MÓDULO DE RELATÓRIOS E GRÁFICOS
' Descrição: Geração de relatórios e visualizações
' ========================================

Option Explicit

' ========================================
' FUNÇÃO: Criar Gráfico de Status dos Projetos
' ========================================
Sub CriarGraficoStatusProjetos()
    Dim ws As Worksheet
    Dim wsDash As Worksheet
    Dim ultimaLinha As Long
    Dim rngDados As Range
    Dim cht As ChartObject
    Dim i As Long
    Dim statusPlanejamento As Long, statusAndamento As Long
    Dim statusPausado As Long, statusCompleto As Long, statusCancelado As Long
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    ' Remover gráfico existente
    On Error Resume Next
    wsDash.ChartObjects("GraficoStatus").Delete
    On Error GoTo 0
    
    ' Contar status
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 2 Then Exit Sub
    
    For i = 2 To ultimaLinha
        Select Case ws.Cells(i, 6).Value
            Case "Planejamento": statusPlanejamento = statusPlanejamento + 1
            Case "Em Andamento": statusAndamento = statusAndamento + 1
            Case "Pausado": statusPausado = statusPausado + 1
            Case "Completo": statusCompleto = statusCompleto + 1
            Case "Cancelado": statusCancelado = statusCancelado + 1
        End Select
    Next i
    
    ' Criar dados para o gráfico
    With wsDash
        .Range("F5").Value = "Status"
        .Range("G5").Value = "Quantidade"
        .Range("F6").Value = "Planejamento"
        .Range("G6").Value = statusPlanejamento
        .Range("F7").Value = "Em Andamento"
        .Range("G7").Value = statusAndamento
        .Range("F8").Value = "Pausado"
        .Range("G8").Value = statusPausado
        .Range("F9").Value = "Completo"
        .Range("G9").Value = statusCompleto
        .Range("F10").Value = "Cancelado"
        .Range("G10").Value = statusCancelado
        
        ' Criar gráfico de pizza
        Set cht = .ChartObjects.Add(Left:=.Range("F12").Left, _
                                    Top:=.Range("F12").Top, _
                                    Width:=350, _
                                    Height:=250)
        
        With cht.Chart
            .SetSourceData Source:=wsDash.Range("F6:G10")
            .ChartType = xlPie
            .HasTitle = True
            .ChartTitle.Text = "Status dos Projetos"
            
            ' Formatação
            .ChartTitle.Font.Size = 12
            .ChartTitle.Font.Bold = True
            
            ' Mostrar valores e percentuais
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowValue = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
            .SeriesCollection(1).DataLabels.Position = xlLabelPositionBestFit
        End With
        
        cht.Name = "GraficoStatus"
    End With
End Sub

' ========================================
' FUNÇÃO: Criar Gráfico de Prioridade das Tarefas
' ========================================
Sub CriarGraficoPrioridadeTarefas()
    Dim ws As Worksheet
    Dim wsDash As Worksheet
    Dim ultimaLinha As Long
    Dim cht As ChartObject
    Dim i As Long
    Dim prioAlta As Long, prioMedia As Long, prioBaixa As Long, prioCritica As Long
    
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    ' Remover gráfico existente
    On Error Resume Next
    wsDash.ChartObjects("GraficoPrioridade").Delete
    On Error GoTo 0
    
    ' Contar prioridades
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 2 Then Exit Sub
    
    For i = 2 To ultimaLinha
        Select Case ws.Cells(i, 8).Value
            Case "Crítica": prioCritica = prioCritica + 1
            Case "Alta": prioAlta = prioAlta + 1
            Case "Média": prioMedia = prioMedia + 1
            Case "Baixa": prioBaixa = prioBaixa + 1
        End Select
    Next i
    
    ' Criar dados para o gráfico
    With wsDash
        .Range("I5").Value = "Prioridade"
        .Range("J5").Value = "Quantidade"
        .Range("I6").Value = "Crítica"
        .Range("J6").Value = prioCritica
        .Range("I7").Value = "Alta"
        .Range("J7").Value = prioAlta
        .Range("I8").Value = "Média"
        .Range("J8").Value = prioMedia
        .Range("I9").Value = "Baixa"
        .Range("J9").Value = prioBaixa
        
        ' Criar gráfico de barras
        Set cht = .ChartObjects.Add(Left:=.Range("I12").Left, _
                                    Top:=.Range("I12").Top, _
                                    Width:=350, _
                                    Height:=250)
        
        With cht.Chart
            .SetSourceData Source:=wsDash.Range("I6:J9")
            .ChartType = xlBarClustered
            .HasTitle = True
            .ChartTitle.Text = "Tarefas por Prioridade"
            
            ' Formatação
            .ChartTitle.Font.Size = 12
            .ChartTitle.Font.Bold = True
            
            ' Remover legenda
            .HasLegend = False
            
            ' Mostrar valores
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
        End With
        
        cht.Name = "GraficoPrioridade"
    End With
End Sub

' ========================================
' FUNÇÃO: Criar Gráfico de Timeline de Projetos
' ========================================
Sub CriarGraficoTimelineProjetos()
    Dim ws As Worksheet
    Dim wsDash As Worksheet
    Dim ultimaLinha As Long
    Dim cht As ChartObject
    Dim i As Long
    Dim linha As Long
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    ' Remover gráfico existente
    On Error Resume Next
    wsDash.ChartObjects("GraficoTimeline").Delete
    On Error GoTo 0
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha < 2 Then Exit Sub
    
    ' Preparar dados
    With wsDash
        .Range("B10").Value = "CRONOGRAMA DE PROJETOS"
        .Range("B10:D10").Merge
        .Range("B10").Font.Bold = True
        
        .Range("B11").Value = "Projeto"
        .Range("C11").Value = "Início"
        .Range("D11").Value = "Fim"
        .Range("E11").Value = "Duração"
        
        linha = 12
        For i = 2 To ultimaLinha
            If ws.Cells(i, 6).Value <> "Cancelado" Then
                .Cells(linha, 2).Value = ws.Cells(i, 2).Value
                .Cells(linha, 3).Value = ws.Cells(i, 4).Value
                .Cells(linha, 4).Value = ws.Cells(i, 5).Value
                .Cells(linha, 5).Value = ws.Cells(i, 5).Value - ws.Cells(i, 4).Value
                linha = linha + 1
            End If
        Next i
        
        If linha > 12 Then
            ' Criar gráfico de Gantt simplificado
            Set cht = .ChartObjects.Add(Left:=.Range("B" & linha + 2).Left, _
                                        Top:=.Range("B" & linha + 2).Top, _
                                        Width:=500, _
                                        Height:=300)
            
            With cht.Chart
                .SetSourceData Source:=wsDash.Range("B11:E" & linha - 1)
                .ChartType = xlBarStacked
                .HasTitle = True
                .ChartTitle.Text = "Cronograma de Projetos"
                
                ' Formatação
                .ChartTitle.Font.Size = 12
                .ChartTitle.Font.Bold = True
            End With
            
            cht.Name = "GraficoTimeline"
        End If
    End With
End Sub

' ========================================
' FUNÇÃO: Gerar Relatório Completo
' ========================================
Sub GerarRelatorioCompleto()
    On Error GoTo TratarErro
    
    Application.ScreenUpdating = False
    
    ' Atualizar dados do dashboard
    Call AtualizarDashboard
    
    ' Criar todos os gráficos
    Call CriarGraficoStatusProjetos
    Call CriarGraficoPrioridadeTarefas
    Call CriarGraficoTimelineProjetos
    
    ' Adicionar análises
    Call GerarAnalisePerformance
    
    Application.ScreenUpdating = True
    
    ' Ativar planilha Dashboard
    ThisWorkbook.Worksheets("Dashboard").Activate
    
    MsgBox "Relatório gerado com sucesso!", vbInformation
    
    Exit Sub
    
TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao gerar relatório: " & Err.Description, vbCritical
End Sub

' ========================================
' FUNÇÃO: Análise de Performance
' ========================================
Private Sub GerarAnalisePerformance()
    Dim wsProjetos As Worksheet, wsTarefas As Worksheet, wsDash As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim totalHorasEstimadas As Double, totalHorasReais As Double
    Dim tarefasAtrasadas As Long, tarefasNoPrazo As Long
    Dim hoje As Date
    
    Set wsProjetos = ThisWorkbook.Worksheets("Projetos")
    Set wsTarefas = ThisWorkbook.Worksheets("Tarefas")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    hoje = Date
    
    ' Análise de horas
    ultimaLinha = wsTarefas.Cells(wsTarefas.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha > 1 Then
        For i = 2 To ultimaLinha
            totalHorasEstimadas = totalHorasEstimadas + wsTarefas.Cells(i, 10).Value
            totalHorasReais = totalHorasReais + wsTarefas.Cells(i, 11).Value
            
            ' Verificar atrasos
            If wsTarefas.Cells(i, 7).Value <> "Completa" And _
               wsTarefas.Cells(i, 6).Value < hoje Then
                tarefasAtrasadas = tarefasAtrasadas + 1
            Else
                tarefasNoPrazo = tarefasNoPrazo + 1
            End If
        Next i
    End If
    
    ' Exibir análise
    With wsDash
        .Range("B24").Value = "ANÁLISE DE PERFORMANCE"
        .Range("B24:D24").Merge
        .Range("B24").Font.Bold = True
        .Range("B24").Interior.Color = COR_HEADER
        .Range("B24").Font.Color = vbWhite
        
        .Range("B25").Value = "Total Horas Estimadas:"
        .Range("C25").Value = totalHorasEstimadas
        .Range("C25").NumberFormat = "0.0"
        
        .Range("B26").Value = "Total Horas Reais:"
        .Range("C26").Value = totalHorasReais
        .Range("C26").NumberFormat = "0.0"
        
        .Range("B27").Value = "Variação:"
        .Range("C27").Value = totalHorasReais - totalHorasEstimadas
        .Range("C27").NumberFormat = "0.0"
        
        If totalHorasEstimadas > 0 Then
            .Range("B28").Value = "% Variação:"
            .Range("C28").Value = (totalHorasReais - totalHorasEstimadas) / totalHorasEstimadas
            .Range("C28").NumberFormat = "0.0%"
        End If
        
        .Range("B30").Value = "Tarefas no Prazo:"
        .Range("C30").Value = tarefasNoPrazo
        
        .Range("B31").Value = "Tarefas Atrasadas:"
        .Range("C31").Value = tarefasAtrasadas
        .Range("C31").Font.Color = vbRed
    End With
End Sub

' ========================================
' FUNÇÃO: Exportar Dashboard para PDF
' ========================================
Sub ExportarDashboardPDF()
    Dim caminhoArquivo As String
    
    On Error GoTo TratarErro
    
    ' Atualizar relatório antes de exportar
    Call GerarRelatorioCompleto
    
    ' Definir caminho do arquivo
    caminhoArquivo = ThisWorkbook.Path & "\Dashboard_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"
    
    ' Exportar para PDF
    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=caminhoArquivo, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    MsgBox "Dashboard exportado para PDF com sucesso!" & vbCrLf & caminhoArquivo, vbInformation
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao exportar PDF: " & Err.Description, vbCritical
End Sub

' ========================================
' FUNÇÃO: Relatório de Projetos por Cliente
' ========================================
Sub RelatorioProjetosPorCliente()
    Dim ws As Worksheet, wsRelatorio As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim clientes As Collection
    Dim cliente As Variant
    Dim linha As Long
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    
    ' Criar ou limpar planilha de relatório
    On Error Resume Next
    Set wsRelatorio = ThisWorkbook.Worksheets("Rel_Clientes")
    On Error GoTo 0
    
    If wsRelatorio Is Nothing Then
        Set wsRelatorio = ThisWorkbook.Worksheets.Add
        wsRelatorio.Name = "Rel_Clientes"
    Else
        wsRelatorio.Cells.Clear
    End If
    
    ' Obter lista única de clientes
    Set clientes = New Collection
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    On Error Resume Next
    For i = 2 To ultimaLinha
        clientes.Add ws.Cells(i, 3).Value, CStr(ws.Cells(i, 3).Value)
    Next i
    On Error GoTo 0
    
    ' Criar cabeçalho
    With wsRelatorio
        .Range("A1").Value = "RELATÓRIO DE PROJETOS POR CLIENTE"
        .Range("A1:F1").Merge
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Cliente"
        .Range("B3").Value = "Total Projetos"
        .Range("C3").Value = "Ativos"
        .Range("D3").Value = "Concluídos"
        .Range("E3").Value = "Orçamento Total"
        .Range("F3").Value = "Progresso Médio"
        
        .Range("A3:F3").Font.Bold = True
        .Range("A3:F3").Interior.Color = COR_HEADER
        .Range("A3:F3").Font.Color = vbWhite
        
        ' Preencher dados
        linha = 4
        For Each cliente In clientes
            .Cells(linha, 1).Value = cliente
            .Cells(linha, 2).Value = Application.WorksheetFunction.CountIf(ws.Columns(3), cliente)
            .Cells(linha, 3).Value = Application.WorksheetFunction.CountIfs(ws.Columns(3), cliente, ws.Columns(6), "Em Andamento")
            .Cells(linha, 4).Value = Application.WorksheetFunction.CountIfs(ws.Columns(3), cliente, ws.Columns(6), "Completo")
            .Cells(linha, 5).Value = Application.WorksheetFunction.SumIf(ws.Columns(3), cliente, ws.Columns(8))
            .Cells(linha, 6).Value = Application.WorksheetFunction.AverageIf(ws.Columns(3), cliente, ws.Columns(7))
            
            linha = linha + 1
        Next cliente
        
        ' Formatação
        .Columns("E:E").NumberFormat = "R$ #,##0.00"
        .Columns("F:F").NumberFormat = "0.0%"
        .Columns("A:F").AutoFit
    End With
    
    wsRelatorio.Activate
    MsgBox "Relatório por cliente gerado com sucesso!", vbInformation
End Sub

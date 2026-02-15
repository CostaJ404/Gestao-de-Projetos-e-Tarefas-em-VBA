Attribute VB_Name = "modPrincipal"
' ========================================
' SISTEMA DE GESTÃO DE PROJETOS E TAREFAS
' Autor: [Seu Nome]
' Data: Fevereiro 2026
' Versão: 1.0
' ========================================

Option Explicit

' ========== CONSTANTES DO SISTEMA ==========
Public Const VERSAO_SISTEMA As String = "1.0"
Public Const COR_HEADER As Long = 5287936      ' Verde escuro
Public Const COR_COMPLETA As Long = 5287936    ' Verde
Public Const COR_ANDAMENTO As Long = 49407     ' Amarelo
Public Const COR_PENDENTE As Long = 255        ' Vermelho

' ========== VARIÁVEIS GLOBAIS ==========
Public ProjetoAtual As Long
Public TarefaAtual As Long

' ========================================
' FUNÇÃO: Inicializar Sistema
' Descrição: Configura o ambiente inicial do sistema
' ========================================
Sub InicializarSistema()
    On Error GoTo TratarErro
    
    Application.ScreenUpdating = False
    
    ' Verificar se as planilhas existem
    Call VerificarPlanilhas
    
    ' Formatar planilhas
    Call FormatarPlanilhaProjetos
    Call FormatarPlanilhaTarefas
    Call FormatarPlanilhaDashboard
    
    ' Atualizar dashboard
    Call AtualizarDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "Sistema inicializado com sucesso!" & vbCrLf & _
           "Versão: " & VERSAO_SISTEMA, vbInformation, "Sistema de Gestão"
    
    Exit Sub
    
TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao inicializar sistema: " & Err.Description, vbCritical
End Sub

' ========================================
' FUNÇÃO: Verificar Planilhas
' Descrição: Cria planilhas necessárias se não existirem
' ========================================
Private Sub VerificarPlanilhas()
    Dim ws As Worksheet
    Dim planilhasNecessarias As Variant
    Dim i As Integer
    
    planilhasNecessarias = Array("Projetos", "Tarefas", "Dashboard", "Equipe")
    
    For i = LBound(planilhasNecessarias) To UBound(planilhasNecessarias)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(planilhasNecessarias(i))
        On Error GoTo 0
        
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = planilhasNecessarias(i)
        End If
        Set ws = Nothing
    Next i
End Sub

' ========================================
' FUNÇÃO: Formatar Planilha Projetos
' Descrição: Configura cabeçalhos e formatação
' ========================================
Private Sub FormatarPlanilhaProjetos()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Projetos")
    
    With ws
        .Cells.Clear
        
        ' Cabeçalhos
        .Range("A1").Value = "ID"
        .Range("B1").Value = "Nome do Projeto"
        .Range("C1").Value = "Cliente"
        .Range("D1").Value = "Data Início"
        .Range("E1").Value = "Data Fim"
        .Range("F1").Value = "Status"
        .Range("G1").Value = "Progresso (%)"
        .Range("H1").Value = "Orçamento"
        .Range("I1").Value = "Gerente"
        .Range("J1").Value = "Descrição"
        
        ' Formatação do cabeçalho
        With .Range("A1:J1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = COR_HEADER
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' Largura das colunas
        .Columns("A:A").ColumnWidth = 8
        .Columns("B:B").ColumnWidth = 25
        .Columns("C:C").ColumnWidth = 20
        .Columns("D:E").ColumnWidth = 12
        .Columns("F:F").ColumnWidth = 12
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 15
        .Columns("I:I").ColumnWidth = 18
        .Columns("J:J").ColumnWidth = 35
        
        ' Formato de número
        .Columns("G:G").NumberFormat = "0%"
        .Columns("H:H").NumberFormat = "R$ #,##0.00"
        .Columns("D:E").NumberFormat = "dd/mm/yyyy"
        
        ' Congelar painéis
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

' ========================================
' FUNÇÃO: Formatar Planilha Tarefas
' Descrição: Configura cabeçalhos e formatação
' ========================================
Private Sub FormatarPlanilhaTarefas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    
    With ws
        .Cells.Clear
        
        ' Cabeçalhos
        .Range("A1").Value = "ID"
        .Range("B1").Value = "ID Projeto"
        .Range("C1").Value = "Tarefa"
        .Range("D1").Value = "Responsável"
        .Range("E1").Value = "Data Início"
        .Range("F1").Value = "Data Fim"
        .Range("G1").Value = "Status"
        .Range("H1").Value = "Prioridade"
        .Range("I1").Value = "Progresso (%)"
        .Range("J1").Value = "Horas Est."
        .Range("K1").Value = "Horas Real"
        .Range("L1").Value = "Observações"
        
        ' Formatação do cabeçalho
        With .Range("A1:L1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = COR_HEADER
            .HorizontalAlignment = xlCenter
        End With
        
        ' Largura das colunas
        .Columns("A:B").ColumnWidth = 8
        .Columns("C:C").ColumnWidth = 30
        .Columns("D:D").ColumnWidth = 18
        .Columns("E:F").ColumnWidth = 12
        .Columns("G:H").ColumnWidth = 12
        .Columns("I:I").ColumnWidth = 12
        .Columns("J:K").ColumnWidth = 10
        .Columns("L:L").ColumnWidth = 35
        
        ' Formato de número
        .Columns("I:I").NumberFormat = "0%"
        .Columns("E:F").NumberFormat = "dd/mm/yyyy"
        
        ' Congelar painéis
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

' ========================================
' FUNÇÃO: Formatar Planilha Dashboard
' Descrição: Cria estrutura do painel de controle
' ========================================
Private Sub FormatarPlanilhaDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    
    With ws
        .Cells.Clear
        
        ' Título
        .Range("B2").Value = "PAINEL DE CONTROLE - GESTÃO DE PROJETOS"
        .Range("B2:H2").Merge
        With .Range("B2")
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = COR_HEADER
            .HorizontalAlignment = xlCenter
        End With
        
        ' Seção de indicadores
        .Range("B4").Value = "INDICADORES GERAIS"
        .Range("B4:D4").Merge
        .Range("B4").Font.Bold = True
        
        .Range("B5").Value = "Total de Projetos:"
        .Range("B6").Value = "Projetos Ativos:"
        .Range("B7").Value = "Tarefas Pendentes:"
        .Range("B8").Value = "Taxa de Conclusão:"
        
        ' Formatação
        .Range("B5:B8").Font.Bold = True
        .Range("C5:C8").NumberFormat = "0"
        .Range("C8").NumberFormat = "0.0%"
        
        ' Largura das colunas
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 15
    End With
End Sub

' ========================================
' FUNÇÃO: Atualizar Dashboard
' Descrição: Calcula e atualiza os indicadores
' ========================================
Sub AtualizarDashboard()
    Dim wsProjetos As Worksheet, wsTarefas As Worksheet, wsDash As Worksheet
    Dim totalProjetos As Long, projetosAtivos As Long
    Dim tarefasPendentes As Long, totalTarefas As Long
    Dim ultimaLinha As Long
    
    Set wsProjetos = ThisWorkbook.Worksheets("Projetos")
    Set wsTarefas = ThisWorkbook.Worksheets("Tarefas")
    Set wsDash = ThisWorkbook.Worksheets("Dashboard")
    
    ' Calcular total de projetos
    ultimaLinha = wsProjetos.Cells(wsProjetos.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha > 1 Then
        totalProjetos = ultimaLinha - 1
    Else
        totalProjetos = 0
    End If
    
    ' Calcular projetos ativos
    If totalProjetos > 0 Then
        projetosAtivos = Application.WorksheetFunction.CountIf(wsProjetos.Range("F2:F" & ultimaLinha), "Em Andamento")
    End If
    
    ' Calcular tarefas pendentes
    ultimaLinha = wsTarefas.Cells(wsTarefas.Rows.Count, 1).End(xlUp).Row
    If ultimaLinha > 1 Then
        totalTarefas = ultimaLinha - 1
        tarefasPendentes = Application.WorksheetFunction.CountIf(wsTarefas.Range("G2:G" & ultimaLinha), "Pendente")
    End If
    
    ' Atualizar valores no dashboard
    With wsDash
        .Range("C5").Value = totalProjetos
        .Range("C6").Value = projetosAtivos
        .Range("C7").Value = tarefasPendentes
        
        If totalTarefas > 0 Then
            .Range("C8").Value = (totalTarefas - tarefasPendentes) / totalTarefas
        Else
            .Range("C8").Value = 0
        End If
    End With
End Sub

' ========================================
' FUNÇÃO: Próximo ID
' Descrição: Retorna o próximo ID disponível
' ========================================
Function ProximoID(nomePlanilha As String) As Long
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    
    Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha = 1 Then
        ProximoID = 1
    Else
        ProximoID = ws.Cells(ultimaLinha, 1).Value + 1
    End If
End Function

' ========================================
' FUNÇÃO: Validar Data
' Descrição: Verifica se a data é válida
' ========================================
Function ValidarData(dataTexto As String) As Boolean
    On Error GoTo DataInvalida
    
    If IsDate(dataTexto) Then
        ValidarData = True
    Else
        ValidarData = False
    End If
    
    Exit Function
    
DataInvalida:
    ValidarData = False
End Function

' ========================================
' FUNÇÃO: Exportar Relatório
' Descrição: Gera relatório em nova pasta de trabalho
' ========================================
Sub ExportarRelatorio()
    Dim wbNovo As Workbook
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim caminhoArquivo As String
    
    On Error GoTo TratarErro
    
    ' Criar nova pasta de trabalho
    Set wbNovo = Workbooks.Add
    
    ' Copiar dados de Projetos
    Set wsOrigem = ThisWorkbook.Worksheets("Projetos")
    Set wsDestino = wbNovo.Worksheets(1)
    wsDestino.Name = "Projetos"
    wsOrigem.UsedRange.Copy wsDestino.Range("A1")
    
    ' Copiar dados de Tarefas
    Set wsOrigem = ThisWorkbook.Worksheets("Tarefas")
    Set wsDestino = wbNovo.Worksheets.Add(After:=wbNovo.Worksheets(wbNovo.Worksheets.Count))
    wsDestino.Name = "Tarefas"
    wsOrigem.UsedRange.Copy wsDestino.Range("A1")
    
    ' Salvar arquivo
    caminhoArquivo = ThisWorkbook.Path & "\Relatorio_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    wbNovo.SaveAs caminhoArquivo
    wbNovo.Close SaveChanges:=False
    
    MsgBox "Relatório exportado com sucesso!" & vbCrLf & caminhoArquivo, vbInformation
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao exportar relatório: " & Err.Description, vbCritical
End Sub

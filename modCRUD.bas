Attribute VB_Name = "modCRUD"
' ========================================
' MÓDULO CRUD - PROJETOS E TAREFAS
' Operações de Create, Read, Update, Delete
' ========================================

Option Explicit

' ========================================
' PROJETOS - CREATE
' ========================================
Sub AdicionarProjeto(nome As String, cliente As String, dataInicio As Date, _
                     dataFim As Date, status As String, progresso As Double, _
                     orcamento As Currency, gerente As String, descricao As String)
    
    Dim ws As Worksheet
    Dim proximaLinha As Long
    Dim novoID As Long
    
    On Error GoTo TratarErro
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    
    ' Validações
    If Len(Trim(nome)) = 0 Then
        MsgBox "O nome do projeto é obrigatório!", vbExclamation
        Exit Sub
    End If
    
    If dataFim < dataInicio Then
        MsgBox "A data final não pode ser anterior à data inicial!", vbExclamation
        Exit Sub
    End If
    
    ' Obter próximo ID
    novoID = ProximoID("Projetos")
    proximaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Inserir dados
    With ws
        .Cells(proximaLinha, 1).Value = novoID
        .Cells(proximaLinha, 2).Value = nome
        .Cells(proximaLinha, 3).Value = cliente
        .Cells(proximaLinha, 4).Value = dataInicio
        .Cells(proximaLinha, 5).Value = dataFim
        .Cells(proximaLinha, 6).Value = status
        .Cells(proximaLinha, 7).Value = progresso / 100
        .Cells(proximaLinha, 8).Value = orcamento
        .Cells(proximaLinha, 9).Value = gerente
        .Cells(proximaLinha, 10).Value = descricao
        
        ' Aplicar formatação condicional por status
        Call FormatarLinhaStatus(.Cells(proximaLinha, 6), proximaLinha)
    End With
    
    ' Atualizar dashboard
    Call AtualizarDashboard
    
    MsgBox "Projeto '" & nome & "' cadastrado com sucesso!" & vbCrLf & _
           "ID: " & novoID, vbInformation
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao adicionar projeto: " & Err.Description, vbCritical
End Sub

' ========================================
' PROJETOS - READ
' ========================================
Function BuscarProjeto(idProjeto As Long) As Collection
    Dim ws As Worksheet
    Dim dados As New Collection
    Dim ultimaLinha As Long, i As Long
    Dim encontrado As Boolean
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    encontrado = False
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = idProjeto Then
            dados.Add ws.Cells(i, 1).Value  ' ID
            dados.Add ws.Cells(i, 2).Value  ' Nome
            dados.Add ws.Cells(i, 3).Value  ' Cliente
            dados.Add ws.Cells(i, 4).Value  ' Data Início
            dados.Add ws.Cells(i, 5).Value  ' Data Fim
            dados.Add ws.Cells(i, 6).Value  ' Status
            dados.Add ws.Cells(i, 7).Value  ' Progresso
            dados.Add ws.Cells(i, 8).Value  ' Orçamento
            dados.Add ws.Cells(i, 9).Value  ' Gerente
            dados.Add ws.Cells(i, 10).Value ' Descrição
            encontrado = True
            Exit For
        End If
    Next i
    
    If encontrado Then
        Set BuscarProjeto = dados
    Else
        Set BuscarProjeto = Nothing
    End If
End Function

' ========================================
' PROJETOS - UPDATE
' ========================================
Sub AtualizarProjeto(idProjeto As Long, nome As String, cliente As String, _
                     dataInicio As Date, dataFim As Date, status As String, _
                     progresso As Double, orcamento As Currency, gerente As String, _
                     descricao As String)
    
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim encontrado As Boolean
    
    On Error GoTo TratarErro
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    encontrado = False
    
    ' Validações
    If dataFim < dataInicio Then
        MsgBox "A data final não pode ser anterior à data inicial!", vbExclamation
        Exit Sub
    End If
    
    ' Buscar e atualizar
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = idProjeto Then
            ws.Cells(i, 2).Value = nome
            ws.Cells(i, 3).Value = cliente
            ws.Cells(i, 4).Value = dataInicio
            ws.Cells(i, 5).Value = dataFim
            ws.Cells(i, 6).Value = status
            ws.Cells(i, 7).Value = progresso / 100
            ws.Cells(i, 8).Value = orcamento
            ws.Cells(i, 9).Value = gerente
            ws.Cells(i, 10).Value = descricao
            
            ' Aplicar formatação
            Call FormatarLinhaStatus(ws.Cells(i, 6), i)
            
            encontrado = True
            Exit For
        End If
    Next i
    
    If encontrado Then
        Call AtualizarDashboard
        MsgBox "Projeto atualizado com sucesso!", vbInformation
    Else
        MsgBox "Projeto não encontrado!", vbExclamation
    End If
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao atualizar projeto: " & Err.Description, vbCritical
End Sub

' ========================================
' PROJETOS - DELETE
' ========================================
Sub ExcluirProjeto(idProjeto As Long)
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim resposta As VbMsgBoxResult
    
    On Error GoTo TratarErro
    
    ' Confirmação
    resposta = MsgBox("Tem certeza que deseja excluir este projeto?" & vbCrLf & _
                      "Esta ação não pode ser desfeita!", vbYesNo + vbQuestion)
    
    If resposta = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = idProjeto Then
            ws.Rows(i).Delete
            
            ' Excluir tarefas relacionadas
            Call ExcluirTarefasPorProjeto(idProjeto)
            
            Call AtualizarDashboard
            MsgBox "Projeto excluído com sucesso!", vbInformation
            Exit Sub
        End If
    Next i
    
    MsgBox "Projeto não encontrado!", vbExclamation
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao excluir projeto: " & Err.Description, vbCritical
End Sub

' ========================================
' TAREFAS - CREATE
' ========================================
Sub AdicionarTarefa(idProjeto As Long, tarefa As String, responsavel As String, _
                    dataInicio As Date, dataFim As Date, status As String, _
                    prioridade As String, progresso As Double, horasEst As Double, _
                    horasReal As Double, observacoes As String)
    
    Dim ws As Worksheet
    Dim proximaLinha As Long
    Dim novoID As Long
    
    On Error GoTo TratarErro
    
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    
    ' Validações
    If Len(Trim(tarefa)) = 0 Then
        MsgBox "A descrição da tarefa é obrigatória!", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o projeto existe
    If BuscarProjeto(idProjeto) Is Nothing Then
        MsgBox "Projeto não encontrado!", vbExclamation
        Exit Sub
    End If
    
    novoID = ProximoID("Tarefas")
    proximaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Inserir dados
    With ws
        .Cells(proximaLinha, 1).Value = novoID
        .Cells(proximaLinha, 2).Value = idProjeto
        .Cells(proximaLinha, 3).Value = tarefa
        .Cells(proximaLinha, 4).Value = responsavel
        .Cells(proximaLinha, 5).Value = dataInicio
        .Cells(proximaLinha, 6).Value = dataFim
        .Cells(proximaLinha, 7).Value = status
        .Cells(proximaLinha, 8).Value = prioridade
        .Cells(proximaLinha, 9).Value = progresso / 100
        .Cells(proximaLinha, 10).Value = horasEst
        .Cells(proximaLinha, 11).Value = horasReal
        .Cells(proximaLinha, 12).Value = observacoes
        
        ' Formatação condicional
        Call FormatarLinhaPrioridade(.Cells(proximaLinha, 8), proximaLinha)
    End With
    
    ' Atualizar progresso do projeto
    Call AtualizarProgressoProjeto(idProjeto)
    Call AtualizarDashboard
    
    MsgBox "Tarefa cadastrada com sucesso!" & vbCrLf & "ID: " & novoID, vbInformation
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao adicionar tarefa: " & Err.Description, vbCritical
End Sub

' ========================================
' TAREFAS - DELETE POR PROJETO
' ========================================
Private Sub ExcluirTarefasPorProjeto(idProjeto As Long)
    Dim ws As Worksheet
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    
    For i = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If ws.Cells(i, 2).Value = idProjeto Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

' ========================================
' FUNÇÃO: Atualizar Progresso do Projeto
' ========================================
Private Sub AtualizarProgressoProjeto(idProjeto As Long)
    Dim wsTarefas As Worksheet, wsProjetos As Worksheet
    Dim ultimaLinha As Long, i As Long, j As Long
    Dim totalTarefas As Long, somaProgresso As Double
    Dim mediaProgresso As Double
    
    Set wsTarefas = ThisWorkbook.Worksheets("Tarefas")
    Set wsProjetos = ThisWorkbook.Worksheets("Projetos")
    
    ultimaLinha = wsTarefas.Cells(wsTarefas.Rows.Count, 1).End(xlUp).Row
    totalTarefas = 0
    somaProgresso = 0
    
    ' Calcular média de progresso das tarefas
    For i = 2 To ultimaLinha
        If wsTarefas.Cells(i, 2).Value = idProjeto Then
            totalTarefas = totalTarefas + 1
            somaProgresso = somaProgresso + wsTarefas.Cells(i, 9).Value
        End If
    Next i
    
    If totalTarefas > 0 Then
        mediaProgresso = somaProgresso / totalTarefas
        
        ' Atualizar progresso no projeto
        ultimaLinha = wsProjetos.Cells(wsProjetos.Rows.Count, 1).End(xlUp).Row
        For j = 2 To ultimaLinha
            If wsProjetos.Cells(j, 1).Value = idProjeto Then
                wsProjetos.Cells(j, 7).Value = mediaProgresso
                Exit For
            End If
        Next j
    End If
End Sub

' ========================================
' FUNÇÃO: Formatar Linha por Status
' ========================================
Private Sub FormatarLinhaStatus(celula As Range, linha As Long)
    Dim ws As Worksheet
    Set ws = celula.Worksheet
    
    Select Case celula.Value
        Case "Completo"
            ws.Rows(linha).Interior.Color = RGB(198, 239, 206)
        Case "Em Andamento"
            ws.Rows(linha).Interior.Color = RGB(255, 235, 156)
        Case "Pendente"
            ws.Rows(linha).Interior.Color = RGB(255, 199, 206)
        Case "Cancelado"
            ws.Rows(linha).Interior.Color = RGB(230, 230, 230)
    End Select
End Sub

' ========================================
' FUNÇÃO: Formatar Linha por Prioridade
' ========================================
Private Sub FormatarLinhaPrioridade(celula As Range, linha As Long)
    Dim ws As Worksheet
    Set ws = celula.Worksheet
    
    Select Case celula.Value
        Case "Alta"
            ws.Cells(linha, 8).Interior.Color = RGB(255, 199, 206)
            ws.Cells(linha, 8).Font.Bold = True
        Case "Média"
            ws.Cells(linha, 8).Interior.Color = RGB(255, 235, 156)
        Case "Baixa"
            ws.Cells(linha, 8).Interior.Color = RGB(198, 239, 206)
    End Select
End Sub

' ========================================
' FUNÇÃO: Listar Todos os Projetos
' ========================================
Function ListarProjetos() As Variant
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim dados As Variant
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha > 1 Then
        dados = ws.Range("A2:J" & ultimaLinha).Value
        ListarProjetos = dados
    Else
        ListarProjetos = Array()
    End If
End Function

' ========================================
' FUNÇÃO: Listar Tarefas por Projeto
' ========================================
Function ListarTarefasPorProjeto(idProjeto As Long) As Variant
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long, contador As Long
    Dim dados() As Variant
    
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha > 1 Then
        ReDim dados(1 To ultimaLinha - 1, 1 To 12)
        contador = 0
        
        For i = 2 To ultimaLinha
            If ws.Cells(i, 2).Value = idProjeto Then
                contador = contador + 1
                dados(contador, 1) = ws.Cells(i, 1).Value   ' ID
                dados(contador, 2) = ws.Cells(i, 2).Value   ' ID Projeto
                dados(contador, 3) = ws.Cells(i, 3).Value   ' Tarefa
                dados(contador, 4) = ws.Cells(i, 4).Value   ' Responsável
                dados(contador, 5) = ws.Cells(i, 5).Value   ' Data Início
                dados(contador, 6) = ws.Cells(i, 6).Value   ' Data Fim
                dados(contador, 7) = ws.Cells(i, 7).Value   ' Status
                dados(contador, 8) = ws.Cells(i, 8).Value   ' Prioridade
                dados(contador, 9) = ws.Cells(i, 9).Value   ' Progresso
                dados(contador, 10) = ws.Cells(i, 10).Value ' Horas Est
                dados(contador, 11) = ws.Cells(i, 11).Value ' Horas Real
                dados(contador, 12) = ws.Cells(i, 12).Value ' Observações
            End If
        Next i
        
        If contador > 0 Then
            ListarTarefasPorProjeto = dados
        Else
            ListarTarefasPorProjeto = Array()
        End If
    Else
        ListarTarefasPorProjeto = Array()
    End If
End Function

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTarefa 
   Caption         =   "Gerenciar Tarefas"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   OleObjectBlob   =   "frmTarefa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' USERFORM: Gerenciar Tarefas
' Descrição: Interface completa para CRUD de tarefas
' ========================================

Option Explicit

Private modoEdicao As Boolean
Private idTarefaAtual As Long

' ========================================
' EVENTO: Inicializar UserForm
' ========================================
Private Sub UserForm_Initialize()
    ' Configurar ComboBox de Projetos
    Call CarregarProjetos
    
    ' Configurar ComboBox de Status
    With cmbStatus
        .AddItem "Pendente"
        .AddItem "Em Andamento"
        .AddItem "Aguardando"
        .AddItem "Completa"
        .AddItem "Cancelada"
        .ListIndex = 0
    End With
    
    ' Configurar ComboBox de Prioridade
    With cmbPrioridade
        .AddItem "Baixa"
        .AddItem "Média"
        .AddItem "Alta"
        .AddItem "Crítica"
        .ListIndex = 1
    End With
    
    ' Configurar ListBox de Tarefas
    Call AtualizarListaTarefas
    
    ' Configurar datas padrão
    txtDataInicio.Value = Format(Date, "dd/mm/yyyy")
    txtDataFim.Value = Format(Date + 7, "dd/mm/yyyy")
    
    ' Configurar valores padrão
    txtProgresso.Value = "0"
    txtHorasEst.Value = "8"
    txtHorasReal.Value = "0"
    
    modoEdicao = False
End Sub

' ========================================
' FUNÇÃO: Carregar Projetos no ComboBox
' ========================================
Private Sub CarregarProjetos()
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim itemProjeto As String
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    cmbProjeto.Clear
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha > 1 Then
        For i = 2 To ultimaLinha
            ' Mostrar apenas projetos ativos
            If ws.Cells(i, 6).Value <> "Completo" And ws.Cells(i, 6).Value <> "Cancelado" Then
                itemProjeto = ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value
                cmbProjeto.AddItem itemProjeto
            End If
        Next i
    End If
    
    If cmbProjeto.ListCount > 0 Then
        cmbProjeto.ListIndex = 0
    End If
End Sub

' ========================================
' FUNÇÃO: Atualizar Lista de Tarefas
' ========================================
Private Sub AtualizarListaTarefas()
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim itemLista As String
    Dim idProjetoFiltro As Long
    
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    lstTarefas.Clear
    
    ' Obter filtro de projeto se selecionado
    idProjetoFiltro = 0
    If cmbProjeto.ListIndex <> -1 Then
        idProjetoFiltro = CLng(Split(cmbProjeto.Value, " - ")(0))
    End If
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha > 1 Then
        For i = 2 To ultimaLinha
            ' Filtrar por projeto se necessário
            If idProjetoFiltro = 0 Or ws.Cells(i, 2).Value = idProjetoFiltro Then
                itemLista = ws.Cells(i, 1).Value & " - " & _
                           ws.Cells(i, 3).Value & " | " & _
                           ws.Cells(i, 4).Value & " | " & _
                           ws.Cells(i, 8).Value & " | " & _
                           ws.Cells(i, 7).Value
                lstTarefas.AddItem itemLista
            End If
        Next i
    End If
End Sub

' ========================================
' BOTÃO: Nova Tarefa
' ========================================
Private Sub btnNovo_Click()
    Call LimparFormulario
    modoEdicao = False
    txtTarefa.SetFocus
End Sub

' ========================================
' BOTÃO: Salvar Tarefa
' ========================================
Private Sub btnSalvar_Click()
    Dim idProjeto As Long
    
    On Error GoTo TratarErro
    
    ' Validações
    If cmbProjeto.ListIndex = -1 Then
        MsgBox "Selecione um projeto!", vbExclamation
        cmbProjeto.SetFocus
        Exit Sub
    End If
    
    If Trim(txtTarefa.Value) = "" Then
        MsgBox "A descrição da tarefa é obrigatória!", vbExclamation
        txtTarefa.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDataInicio.Value) Then
        MsgBox "Data de início inválida!", vbExclamation
        txtDataInicio.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(txtDataFim.Value) Then
        MsgBox "Data final inválida!", vbExclamation
        txtDataFim.SetFocus
        Exit Sub
    End If
    
    ' Obter ID do projeto
    idProjeto = CLng(Split(cmbProjeto.Value, " - ")(0))
    
    ' Salvar tarefa
    If modoEdicao Then
        ' Implementar atualização de tarefa
        MsgBox "Função de edição em desenvolvimento!", vbInformation
    Else
        Call AdicionarTarefa(idProjeto, _
                            txtTarefa.Value, _
                            txtResponsavel.Value, _
                            CDate(txtDataInicio.Value), _
                            CDate(txtDataFim.Value), _
                            cmbStatus.Value, _
                            cmbPrioridade.Value, _
                            CDbl(txtProgresso.Value), _
                            CDbl(txtHorasEst.Value), _
                            CDbl(txtHorasReal.Value), _
                            txtObservacoes.Value)
    End If
    
    Call LimparFormulario
    Call AtualizarListaTarefas
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

' ========================================
' BOTÃO: Fechar Formulário
' ========================================
Private Sub btnFechar_Click()
    Unload Me
End Sub

' ========================================
' FUNÇÃO: Limpar Formulário
' ========================================
Private Sub LimparFormulario()
    txtTarefa.Value = ""
    txtResponsavel.Value = ""
    txtDataInicio.Value = Format(Date, "dd/mm/yyyy")
    txtDataFim.Value = Format(Date + 7, "dd/mm/yyyy")
    cmbStatus.ListIndex = 0
    cmbPrioridade.ListIndex = 1
    txtProgresso.Value = "0"
    txtHorasEst.Value = "8"
    txtHorasReal.Value = "0"
    txtObservacoes.Value = ""
    
    modoEdicao = False
    idTarefaAtual = 0
End Sub

' ========================================
' EVENTO: Mudança no ComboBox de Projeto
' ========================================
Private Sub cmbProjeto_Change()
    Call AtualizarListaTarefas
End Sub

' ========================================
' EVENTO: Mudança no progresso
' ========================================
Private Sub txtProgresso_Change()
    On Error Resume Next
    If IsNumeric(txtProgresso.Value) Then
        Dim valor As Integer
        valor = CInt(txtProgresso.Value)
        If valor < 0 Then txtProgresso.Value = "0"
        If valor > 100 Then txtProgresso.Value = "100"
        
        ' Atualizar status automaticamente
        If valor = 100 Then
            cmbStatus.Value = "Completa"
        ElseIf valor > 0 Then
            If cmbStatus.Value = "Pendente" Then
                cmbStatus.Value = "Em Andamento"
            End If
        End If
    End If
End Sub

' ========================================
' BOTÃO: Filtrar Tarefas
' ========================================
Private Sub btnFiltrar_Click()
    Call AtualizarListaTarefas
End Sub

' ========================================
' BOTÃO: Ver Todas as Tarefas
' ========================================
Private Sub btnVerTodas_Click()
    cmbProjeto.ListIndex = -1
    Call AtualizarListaTarefas
End Sub

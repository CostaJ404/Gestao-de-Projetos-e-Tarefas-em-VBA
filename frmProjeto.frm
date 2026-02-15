VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjeto 
   Caption         =   "Gerenciar Projetos"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "frmProjeto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ========================================
' USERFORM: Gerenciar Projetos
' Descrição: Interface completa para CRUD de projetos
' ========================================

Option Explicit

Private modoEdicao As Boolean
Private idProjetoAtual As Long

' ========================================
' EVENTO: Inicializar UserForm
' ========================================
Private Sub UserForm_Initialize()
    ' Configurar ComboBox de Status
    With cmbStatus
        .AddItem "Planejamento"
        .AddItem "Em Andamento"
        .AddItem "Pausado"
        .AddItem "Completo"
        .AddItem "Cancelado"
        .ListIndex = 0
    End With
    
    ' Configurar ListBox de Projetos
    Call AtualizarListaProjetos
    
    ' Configurar datas padrão
    txtDataInicio.Value = Format(Date, "dd/mm/yyyy")
    txtDataFim.Value = Format(Date + 30, "dd/mm/yyyy")
    
    ' Configurar barra de progresso
    txtProgresso.Value = "0"
    
    modoEdicao = False
End Sub

' ========================================
' FUNÇÃO: Atualizar Lista de Projetos
' ========================================
Private Sub AtualizarListaProjetos()
    Dim ws As Worksheet
    Dim ultimaLinha As Long, i As Long
    Dim itemLista As String
    
    Set ws = ThisWorkbook.Worksheets("Projetos")
    lstProjetos.Clear
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaLinha > 1 Then
        For i = 2 To ultimaLinha
            itemLista = ws.Cells(i, 1).Value & " - " & _
                       ws.Cells(i, 2).Value & " | " & _
                       ws.Cells(i, 3).Value & " | " & _
                       ws.Cells(i, 6).Value
            lstProjetos.AddItem itemLista
        Next i
    End If
End Sub

' ========================================
' BOTÃO: Novo Projeto
' ========================================
Private Sub btnNovo_Click()
    Call LimparFormulario
    modoEdicao = False
    txtNome.SetFocus
End Sub

' ========================================
' BOTÃO: Salvar Projeto
' ========================================
Private Sub btnSalvar_Click()
    On Error GoTo TratarErro
    
    ' Validações
    If Trim(txtNome.Value) = "" Then
        MsgBox "O nome do projeto é obrigatório!", vbExclamation
        txtNome.SetFocus
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
    
    If Not IsNumeric(txtProgresso.Value) Then
        MsgBox "Progresso inválido!", vbExclamation
        txtProgresso.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtOrcamento.Value) Then
        MsgBox "Orçamento inválido!", vbExclamation
        txtOrcamento.SetFocus
        Exit Sub
    End If
    
    ' Salvar ou Atualizar
    If modoEdicao Then
        Call AtualizarProjeto(idProjetoAtual, _
                             txtNome.Value, _
                             txtCliente.Value, _
                             CDate(txtDataInicio.Value), _
                             CDate(txtDataFim.Value), _
                             cmbStatus.Value, _
                             CDbl(txtProgresso.Value), _
                             CCur(txtOrcamento.Value), _
                             txtGerente.Value, _
                             txtDescricao.Value)
    Else
        Call AdicionarProjeto(txtNome.Value, _
                             txtCliente.Value, _
                             CDate(txtDataInicio.Value), _
                             CDate(txtDataFim.Value), _
                             cmbStatus.Value, _
                             CDbl(txtProgresso.Value), _
                             CCur(txtOrcamento.Value), _
                             txtGerente.Value, _
                             txtDescricao.Value)
    End If
    
    Call LimparFormulario
    Call AtualizarListaProjetos
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao salvar: " & Err.Description, vbCritical
End Sub

' ========================================
' BOTÃO: Editar Projeto
' ========================================
Private Sub btnEditar_Click()
    Dim idProjeto As Long
    Dim dados As Collection
    
    If lstProjetos.ListIndex = -1 Then
        MsgBox "Selecione um projeto na lista!", vbExclamation
        Exit Sub
    End If
    
    ' Obter ID do projeto selecionado
    idProjeto = CLng(Split(lstProjetos.Value, " - ")(0))
    
    ' Buscar dados do projeto
    Set dados = BuscarProjeto(idProjeto)
    
    If Not dados Is Nothing Then
        idProjetoAtual = dados(1)
        txtNome.Value = dados(2)
        txtCliente.Value = dados(3)
        txtDataInicio.Value = Format(dados(4), "dd/mm/yyyy")
        txtDataFim.Value = Format(dados(5), "dd/mm/yyyy")
        cmbStatus.Value = dados(6)
        txtProgresso.Value = dados(7) * 100
        txtOrcamento.Value = dados(8)
        txtGerente.Value = dados(9)
        txtDescricao.Value = dados(10)
        
        modoEdicao = True
        txtNome.SetFocus
    End If
End Sub

' ========================================
' BOTÃO: Excluir Projeto
' ========================================
Private Sub btnExcluir_Click()
    Dim idProjeto As Long
    
    If lstProjetos.ListIndex = -1 Then
        MsgBox "Selecione um projeto na lista!", vbExclamation
        Exit Sub
    End If
    
    ' Obter ID do projeto
    idProjeto = CLng(Split(lstProjetos.Value, " - ")(0))
    
    ' Excluir
    Call ExcluirProjeto(idProjeto)
    
    Call LimparFormulario
    Call AtualizarListaProjetos
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
    txtNome.Value = ""
    txtCliente.Value = ""
    txtDataInicio.Value = Format(Date, "dd/mm/yyyy")
    txtDataFim.Value = Format(Date + 30, "dd/mm/yyyy")
    cmbStatus.ListIndex = 0
    txtProgresso.Value = "0"
    txtOrcamento.Value = ""
    txtGerente.Value = ""
    txtDescricao.Value = ""
    
    modoEdicao = False
    idProjetoAtual = 0
End Sub

' ========================================
' EVENTO: Duplo clique na lista
' ========================================
Private Sub lstProjetos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnEditar_Click
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
    End If
End Sub

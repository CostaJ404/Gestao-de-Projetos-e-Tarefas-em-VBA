Attribute VB_Name = "modMenu"
' ========================================
' MÓDULO DE MENU E ATALHOS
' Descrição: Procedimentos para abrir formulários e atalhos
' ========================================

Option Explicit

' ========================================
' PROCEDIMENTO: Abrir Formulário de Projetos
' ========================================
Sub AbrirFormularioProjetos()
    On Error GoTo TratarErro
    
    ' Verificar se o sistema foi inicializado
    If Not VerificarInicializacao Then
        If MsgBox("O sistema ainda não foi inicializado. Deseja inicializar agora?", _
                  vbYesNo + vbQuestion, "Sistema") = vbYes Then
            Call InicializarSistema
        Else
            Exit Sub
        End If
    End If
    
    ' Abrir formulário
    frmProjeto.Show
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao abrir formulário de projetos: " & Err.Description, vbCritical
End Sub

' ========================================
' PROCEDIMENTO: Abrir Formulário de Tarefas
' ========================================
Sub AbrirFormularioTarefas()
    On Error GoTo TratarErro
    
    ' Verificar se o sistema foi inicializado
    If Not VerificarInicializacao Then
        If MsgBox("O sistema ainda não foi inicializado. Deseja inicializar agora?", _
                  vbYesNo + vbQuestion, "Sistema") = vbYes Then
            Call InicializarSistema
        Else
            Exit Sub
        End If
    End If
    
    ' Abrir formulário
    frmTarefa.Show
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao abrir formulário de tarefas: " & Err.Description, vbCritical
End Sub

' ========================================
' FUNÇÃO: Verificar se Sistema foi Inicializado
' ========================================
Private Function VerificarInicializacao() As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Projetos")
    
    If ws Is Nothing Then
        VerificarInicializacao = False
    Else
        ' Verificar se tem cabeçalhos
        If ws.Range("A1").Value = "ID" Then
            VerificarInicializacao = True
        Else
            VerificarInicializacao = False
        End If
    End If
End Function

' ========================================
' PROCEDIMENTO: Ir para Dashboard
' ========================================
Sub IrParaDashboard()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("Dashboard").Activate
    
    If Err.Number <> 0 Then
        MsgBox "Dashboard não encontrado. Execute 'Inicializar Sistema' primeiro.", vbExclamation
    End If
End Sub

' ========================================
' PROCEDIMENTO: Ir para Projetos
' ========================================
Sub IrParaProjetos()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("Projetos").Activate
    
    If Err.Number <> 0 Then
        MsgBox "Planilha de Projetos não encontrada. Execute 'Inicializar Sistema' primeiro.", vbExclamation
    End If
End Sub

' ========================================
' PROCEDIMENTO: Ir para Tarefas
' ========================================
Sub IrParaTarefas()
    On Error Resume Next
    
    ThisWorkbook.Worksheets("Tarefas").Activate
    
    If Err.Number <> 0 Then
        MsgBox "Planilha de Tarefas não encontrada. Execute 'Inicializar Sistema' primeiro.", vbExclamation
    End If
End Sub

' ========================================
' PROCEDIMENTO: Criar Atalhos de Teclado
' Descrição: Configura atalhos personalizados
' ========================================
Sub ConfigurarAtalhos()
    ' Atalhos podem ser configurados através de:
    ' Arquivo > Opções > Personalizar Faixa de Opções > Atalhos de Teclado
    
    ' Ou criar um evento Workbook_Open para atribuir programaticamente
    ' Exemplo: Application.OnKey "^p", "AbrirFormularioProjetos"
    ' (Ctrl + P para abrir Projetos)
    
    MsgBox "Atalhos disponíveis:" & vbCrLf & vbCrLf & _
           "Use os botões no Menu para:" & vbCrLf & _
           "- Gerenciar Projetos" & vbCrLf & _
           "- Gerenciar Tarefas" & vbCrLf & _
           "- Gerar Relatórios" & vbCrLf & _
           "- Exportar Dashboard", vbInformation, "Atalhos do Sistema"
End Sub

' ========================================
' PROCEDIMENTO: Backup do Sistema
' Descrição: Cria uma cópia de segurança
' ========================================
Sub FazerBackup()
    Dim caminhoBackup As String
    Dim nomeArquivo As String
    
    On Error GoTo TratarErro
    
    ' Definir caminho e nome do backup
    nomeArquivo = "Backup_Sistema_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    caminhoBackup = ThisWorkbook.Path & "\" & nomeArquivo
    
    ' Salvar cópia
    ThisWorkbook.SaveCopyAs caminhoBackup
    
    MsgBox "Backup criado com sucesso!" & vbCrLf & vbCrLf & _
           "Local: " & caminhoBackup, vbInformation, "Backup"
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao criar backup: " & Err.Description, vbCritical
End Sub

' ========================================
' PROCEDIMENTO: Limpar Todos os Dados
' Descrição: Remove todos os dados (mantém estrutura)
' CUIDADO: Esta ação não pode ser desfeita!
' ========================================
Sub LimparTodosDados()
    Dim resposta As VbMsgBoxResult
    Dim ws As Worksheet
    
    On Error GoTo TratarErro
    
    ' Confirmação dupla
    resposta = MsgBox("ATENÇÃO!" & vbCrLf & vbCrLf & _
                      "Esta ação irá APAGAR TODOS OS DADOS do sistema!" & vbCrLf & _
                      "Projetos, Tarefas e Relatórios serão removidos." & vbCrLf & vbCrLf & _
                      "Esta ação NÃO PODE SER DESFEITA!" & vbCrLf & vbCrLf & _
                      "Deseja continuar?", vbYesNo + vbCritical, "ATENÇÃO!")
    
    If resposta = vbNo Then Exit Sub
    
    ' Segunda confirmação
    resposta = MsgBox("Tem ABSOLUTA CERTEZA?" & vbCrLf & _
                      "Todos os dados serão perdidos!", vbYesNo + vbExclamation, "Última Chance")
    
    If resposta = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Limpar dados da planilha Projetos
    Set ws = ThisWorkbook.Worksheets("Projetos")
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Rows("2:" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Delete
    End If
    
    ' Limpar dados da planilha Tarefas
    Set ws = ThisWorkbook.Worksheets("Tarefas")
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Rows("2:" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).Delete
    End If
    
    ' Limpar gráficos do Dashboard
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error Resume Next
    ws.ChartObjects.Delete
    On Error GoTo TratarErro
    
    ' Atualizar dashboard
    Call AtualizarDashboard
    
    Application.ScreenUpdating = True
    
    MsgBox "Todos os dados foram removidos!" & vbCrLf & _
           "A estrutura do sistema foi mantida.", vbInformation, "Limpeza Concluída"
    
    Exit Sub
    
TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao limpar dados: " & Err.Description, vbCritical
End Sub

' ========================================
' PROCEDIMENTO: Sobre o Sistema
' Descrição: Exibe informações sobre o sistema
' ========================================
Sub SobreOSistema()
    Dim mensagem As String
    
    mensagem = "SISTEMA DE GESTÃO DE PROJETOS E TAREFAS" & vbCrLf & _
               String(50, "=") & vbCrLf & vbCrLf & _
               "Versão: " & VERSAO_SISTEMA & vbCrLf & _
               "Desenvolvido em: VBA Excel" & vbCrLf & vbCrLf & _
               "FUNCIONALIDADES:" & vbCrLf & _
               "✓ Gestão completa de projetos" & vbCrLf & _
               "✓ Controle de tarefas e prioridades" & vbCrLf & _
               "✓ Dashboard com indicadores" & vbCrLf & _
               "✓ Relatórios automáticos" & vbCrLf & _
               "✓ Análise de performance" & vbCrLf & _
               "✓ Exportação para PDF" & vbCrLf & vbCrLf & _
               "© 2026 - Todos os direitos reservados"
    
    MsgBox mensagem, vbInformation, "Sobre o Sistema"
End Sub

' ========================================
' PROCEDIMENTO: Ajuda do Sistema
' Descrição: Guia rápido de uso
' ========================================
Sub AjudaSistema()
    Dim mensagem As String
    
    mensagem = "GUIA RÁPIDO DE USO" & vbCrLf & _
               String(50, "=") & vbCrLf & vbCrLf & _
               "1. INICIAR:" & vbCrLf & _
               "   Clique em 'Inicializar Sistema' (apenas na primeira vez)" & vbCrLf & vbCrLf & _
               "2. PROJETOS:" & vbCrLf & _
               "   Clique em 'Gerenciar Projetos' para adicionar/editar" & vbCrLf & vbCrLf & _
               "3. TAREFAS:" & vbCrLf & _
               "   Clique em 'Gerenciar Tarefas' e selecione o projeto" & vbCrLf & vbCrLf & _
               "4. RELATÓRIOS:" & vbCrLf & _
               "   Clique em 'Gerar Relatórios' para atualizar o dashboard" & vbCrLf & vbCrLf & _
               "5. EXPORTAR:" & vbCrLf & _
               "   Clique em 'Exportar Dashboard' para salvar em PDF" & vbCrLf & vbCrLf & _
               "DICA: Use os formulários para gerenciar dados facilmente!"
    
    MsgBox mensagem, vbInformation, "Ajuda do Sistema"
End Sub

' ========================================
' PROCEDIMENTO: Dados de Demonstração
' Descrição: Cria dados de exemplo para teste
' ========================================
Sub CriarDadosDemonstracao()
    Dim resposta As VbMsgBoxResult
    
    resposta = MsgBox("Deseja criar dados de demonstração?" & vbCrLf & vbCrLf & _
                      "Serão criados:" & vbCrLf & _
                      "- 3 projetos de exemplo" & vbCrLf & _
                      "- 6 tarefas de exemplo" & vbCrLf & vbCrLf & _
                      "Ideal para testar o sistema!", vbYesNo + vbQuestion, "Dados de Demonstração")
    
    If resposta = vbNo Then Exit Sub
    
    On Error GoTo TratarErro
    
    ' Projeto 1
    Call AdicionarProjeto("Website Corporativo", "Empresa ABC Ltda", _
                         DateSerial(2026, 2, 1), DateSerial(2026, 4, 1), _
                         "Em Andamento", 45, 75000, "João Silva", _
                         "Desenvolvimento do novo website institucional com design moderno")
    
    ' Projeto 2
    Call AdicionarProjeto("Sistema de Vendas", "Comércio XYZ", _
                         DateSerial(2026, 1, 15), DateSerial(2026, 5, 30), _
                         "Em Andamento", 30, 120000, "Maria Santos", _
                         "Sistema completo de gestão de vendas e estoque")
    
    ' Projeto 3
    Call AdicionarProjeto("App Mobile", "Tech Start Ltda", _
                         DateSerial(2026, 3, 1), DateSerial(2026, 6, 30), _
                         "Planejamento", 10, 95000, "Pedro Costa", _
                         "Aplicativo mobile para iOS e Android")
    
    ' Tarefas do Projeto 1
    Call AdicionarTarefa(1, "Design da Homepage", "Ana Paula", _
                        DateSerial(2026, 2, 1), DateSerial(2026, 2, 15), _
                        "Completa", "Alta", 100, 40, 38, _
                        "Design aprovado pelo cliente")
    
    Call AdicionarTarefa(1, "Desenvolvimento Backend", "Carlos Mendes", _
                        DateSerial(2026, 2, 10), DateSerial(2026, 3, 15), _
                        "Em Andamento", "Alta", 60, 80, 52, _
                        "API REST em desenvolvimento")
    
    ' Tarefas do Projeto 2
    Call AdicionarTarefa(2, "Análise de Requisitos", "Maria Santos", _
                        DateSerial(2026, 1, 15), DateSerial(2026, 1, 30), _
                        "Completa", "Crítica", 100, 40, 45, _
                        "Requisitos levantados e documentados")
    
    Call AdicionarTarefa(2, "Modelagem do Banco de Dados", "Roberto Lima", _
                        DateSerial(2026, 2, 1), DateSerial(2026, 2, 20), _
                        "Em Andamento", "Alta", 70, 60, 48, _
                        "Modelo ER completo")
    
    ' Tarefas do Projeto 3
    Call AdicionarTarefa(3, "Prototipação das Telas", "Juliana Soares", _
                        DateSerial(2026, 3, 1), DateSerial(2026, 3, 10), _
                        "Pendente", "Média", 0, 30, 0, _
                        "Aguardando aprovação do briefing")
    
    Call AdicionarTarefa(3, "Definição da Stack Tecnológica", "Pedro Costa", _
                        DateSerial(2026, 3, 5), DateSerial(2026, 3, 12), _
                        "Pendente", "Alta", 0, 16, 0, _
                        "Avaliar React Native vs Flutter")
    
    ' Gerar relatórios
    Call GerarRelatorioCompleto
    
    MsgBox "Dados de demonstração criados com sucesso!" & vbCrLf & vbCrLf & _
           "✓ 3 projetos criados" & vbCrLf & _
           "✓ 6 tarefas criadas" & vbCrLf & _
           "✓ Dashboard atualizado", vbInformation, "Sucesso!"
    
    Exit Sub
    
TratarErro:
    MsgBox "Erro ao criar dados de demonstração: " & Err.Description, vbCritical
End Sub

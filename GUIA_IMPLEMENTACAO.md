# 📊 SISTEMA DE GESTÃO DE PROJETOS E TAREFAS - VBA EXCEL

## 🎯 Sobre o Projeto

Este é um **sistema completo de gestão de projetos e tarefas** desenvolvido em VBA para Excel, criado para demonstrar habilidades profissionais em automação e desenvolvimento de soluções corporativas.

### ✨ Características Principais

- ✅ **Interface Gráfica Completa** - UserForms profissionais
- ✅ **CRUD Completo** - Create, Read, Update, Delete
- ✅ **Dashboard Interativo** - Indicadores e gráficos em tempo real
- ✅ **Relatórios Automatizados** - Exportação para PDF e Excel
- ✅ **Validações Robustas** - Tratamento de erros e validação de dados
- ✅ **Código Documentado** - Comentários e estrutura profissional
- ✅ **Análise de Performance** - Métricas e KPIs automáticos

---

## 📁 Estrutura do Projeto

### Módulos VBA (4 arquivos .bas)

1. **modPrincipal.bas** - Módulo principal do sistema
   - Inicialização do sistema
   - Formatação das planilhas
   - Atualização do dashboard
   - Funções auxiliares

2. **modCRUD.bas** - Operações de banco de dados
   - Adicionar, buscar, atualizar e excluir projetos
   - Adicionar, buscar, atualizar e excluir tarefas
   - Formatação condicional automática
   - Cálculo de progresso

3. **modRelatorios.bas** - Geração de relatórios e gráficos
   - Gráfico de status dos projetos
   - Gráfico de prioridade das tarefas
   - Timeline de projetos
   - Análise de performance
   - Exportação para PDF
   - Relatórios por cliente

### UserForms (2 arquivos .frm)

4. **frmProjeto.frm** - Formulário de gerenciamento de projetos
   - Cadastro de novos projetos
   - Edição de projetos existentes
   - Exclusão de projetos
   - Listagem e filtros

5. **frmTarefa.frm** - Formulário de gerenciamento de tarefas
   - Cadastro de tarefas vinculadas a projetos
   - Controle de progresso
   - Gestão de prioridades
   - Acompanhamento de horas

---

## 🚀 GUIA DE IMPLEMENTAÇÃO PASSO A PASSO

### PASSO 1: Criar a Pasta de Trabalho

1. Abra o Microsoft Excel
2. Crie uma nova pasta de trabalho
3. Salve como: **"Sistema_Gestao_Projetos.xlsm"** (formato Habilitado para Macros)
4. **IMPORTANTE**: Certifique-se de que o arquivo está salvo no formato `.xlsm`

---

### PASSO 2: Habilitar a Guia Desenvolvedor

1. Vá em **Arquivo** → **Opções**
2. Clique em **Personalizar Faixa de Opções**
3. No lado direito, marque a caixa **Desenvolvedor**
4. Clique em **OK**

---

### PASSO 3: Importar os Módulos VBA

#### 3.1 Abrir o Editor VBA
- Pressione **Alt + F11** ou
- Vá em **Desenvolvedor** → **Visual Basic**

#### 3.2 Importar os Módulos (.bas)

Para cada arquivo `.bas`:

1. No Editor VBA, clique em **Arquivo** → **Importar Arquivo**
2. Navegue até a pasta onde salvou os arquivos
3. Selecione o arquivo e clique em **Abrir**
4. Repita para todos os 4 módulos:
   - modPrincipal.bas
   - modCRUD.bas
   - modRelatorios.bas

**Você verá os módulos aparecerem na janela do Project Explorer à esquerda.**

---

### PASSO 4: Criar os UserForms

#### 4.1 Criar UserForm de Projetos

1. No Editor VBA, clique em **Inserir** → **UserForm**
2. Um novo formulário em branco aparecerá
3. Na janela **Propriedades** (F4), encontre a propriedade **Name**
4. Altere o nome para: **frmProjeto**
5. Altere a propriedade **Caption** para: **Gerenciar Projetos**

#### 4.2 Adicionar Controles ao Formulário de Projetos

Adicione os seguintes controles (da Caixa de Ferramentas):

**Labels e TextBoxes:**
- Label: "Nome do Projeto:" → TextBox: **txtNome**
- Label: "Cliente:" → TextBox: **txtCliente**
- Label: "Data Início:" → TextBox: **txtDataInicio**
- Label: "Data Fim:" → TextBox: **txtDataFim**
- Label: "Orçamento (R$):" → TextBox: **txtOrcamento**
- Label: "Gerente:" → TextBox: **txtGerente**
- Label: "Progresso (%):" → TextBox: **txtProgresso**
- Label: "Descrição:" → TextBox: **txtDescricao** (MultiLine = True)

**ComboBox:**
- Label: "Status:" → ComboBox: **cmbStatus**

**ListBox:**
- Label: "Projetos Cadastrados:" → ListBox: **lstProjetos**

**Botões (CommandButton):**
- **btnNovo** - Caption: "Novo"
- **btnSalvar** - Caption: "Salvar"
- **btnEditar** - Caption: "Editar"
- **btnExcluir** - Caption: "Excluir"
- **btnFechar** - Caption: "Fechar"

#### 4.3 Copiar o Código do UserForm de Projetos

1. Clique duas vezes no formulário para abrir a janela de código
2. **APAGUE** todo o código existente
3. Abra o arquivo **frmProjeto.frm** que você salvou
4. **COPIE TODO O CÓDIGO** (do `Option Explicit` até o final)
5. **COLE** na janela de código do UserForm

#### 4.4 Criar UserForm de Tarefas

Repita o processo:
1. **Inserir** → **UserForm**
2. Name: **frmTarefa**
3. Caption: **Gerenciar Tarefas**

**Controles necessários:**

**ComboBoxes:**
- **cmbProjeto** - Lista de projetos
- **cmbStatus** - Status da tarefa
- **cmbPrioridade** - Prioridade da tarefa

**TextBoxes:**
- **txtTarefa** - Descrição da tarefa
- **txtResponsavel** - Nome do responsável
- **txtDataInicio** - Data de início
- **txtDataFim** - Data final
- **txtProgresso** - Progresso (%)
- **txtHorasEst** - Horas estimadas
- **txtHorasReal** - Horas reais
- **txtObservacoes** - Observações (MultiLine = True)

**ListBox:**
- **lstTarefas** - Lista de tarefas

**Botões:**
- **btnNovo**, **btnSalvar**, **btnFechar**, **btnFiltrar**, **btnVerTodas**

Copie o código do arquivo **frmTarefa.frm**

---

### PASSO 5: Criar o Menu Principal

#### 5.1 Criar uma Planilha de Menu

1. Volte para o Excel (Alt + F11 para sair do VBA)
2. Insira uma nova planilha
3. Renomeie para **"Menu"**
4. Posicione-a como primeira aba

#### 5.2 Formatar o Menu

Crie um design atrativo:

```
Célula B2: "SISTEMA DE GESTÃO DE PROJETOS"
Célula B4: "Bem-vindo ao Sistema de Gestão!"
Célula B6: "Escolha uma opção abaixo:"
```

#### 5.3 Criar Botões de Ação

1. Vá em **Desenvolvedor** → **Inserir** → **Botão (Controle de Formulário)**
2. Desenhe um botão
3. Na caixa de diálogo, atribua a macro correspondente
4. Clique com o botão direito no botão → **Editar Texto**

**Criar 5 botões:**

**Botão 1: "Inicializar Sistema"**
- Macro: `InicializarSistema`

**Botão 2: "Gerenciar Projetos"**
- Macro: Criar uma nova macro:
```vba
Sub AbrirFormularioProjetos()
    frmProjeto.Show
End Sub
```

**Botão 3: "Gerenciar Tarefas"**
- Macro: Criar uma nova macro:
```vba
Sub AbrirFormularioTarefas()
    frmTarefa.Show
End Sub
```

**Botão 4: "Gerar Relatórios"**
- Macro: `GerarRelatorioCompleto`

**Botão 5: "Exportar Dashboard (PDF)"**
- Macro: `ExportarDashboardPDF`

---

### PASSO 6: Inicializar o Sistema

1. Vá para a planilha **Menu**
2. Clique no botão **"Inicializar Sistema"**
3. O sistema criará automaticamente as planilhas:
   - Projetos
   - Tarefas
   - Dashboard
   - Equipe

4. Todas as planilhas serão formatadas automaticamente

---

### PASSO 7: Testar o Sistema

#### Teste 1: Adicionar um Projeto
1. Clique em **"Gerenciar Projetos"**
2. Preencha os dados:
   - Nome: "Website Corporativo"
   - Cliente: "Empresa ABC"
   - Data Início: 01/02/2026
   - Data Fim: 01/04/2026
   - Status: Em Andamento
   - Progresso: 30
   - Orçamento: 50000
   - Gerente: João Silva
3. Clique em **Salvar**

#### Teste 2: Adicionar Tarefas
1. Clique em **"Gerenciar Tarefas"**
2. Selecione o projeto criado
3. Adicione tarefas:
   - Tarefa: "Design do Layout"
   - Responsável: Maria Santos
   - Prioridade: Alta
   - Status: Em Andamento
4. Clique em **Salvar**

#### Teste 3: Gerar Relatórios
1. Clique em **"Gerar Relatórios"**
2. Verifique o Dashboard atualizado com:
   - Gráfico de status
   - Gráfico de prioridades
   - Análise de performance

---

## 🎨 CUSTOMIZAÇÕES SUGERIDAS

### Personalizar Cores

No módulo `modPrincipal.bas`, altere as constantes:

```vba
Public Const COR_HEADER As Long = 5287936      ' Verde escuro
Public Const COR_COMPLETA As Long = 5287936    ' Verde
Public Const COR_ANDAMENTO As Long = 49407     ' Amarelo
Public Const COR_PENDENTE As Long = 255        ' Vermelho
```

### Adicionar Logo da Empresa

1. Vá para a planilha **Menu**
2. Insira uma imagem do logo
3. Posicione e redimensione conforme necessário

---

## 📊 FUNCIONALIDADES AVANÇADAS

### 1. Validação de Dados
- Datas não podem ser retroativas
- Progresso limitado entre 0-100%
- Orçamento deve ser numérico
- Campos obrigatórios validados

### 2. Formatação Condicional Automática
- Projetos **concluídos**: Verde
- Projetos **em andamento**: Amarelo
- Projetos **pendentes**: Vermelho
- Prioridade **alta**: Destaque vermelho

### 3. Cálculos Automáticos
- Progresso do projeto calculado pela média das tarefas
- Total de horas estimadas vs. reais
- Identificação automática de tarefas atrasadas

### 4. Relatórios
- Status dos projetos (gráfico de pizza)
- Prioridade das tarefas (gráfico de barras)
- Análise de performance
- Relatório por cliente
- Exportação para PDF

## 🔒 SEGURANÇA E MACROS

### Habilitar Macros
1. Vá em **Arquivo** → **Opções**
2. **Central de Confiabilidade** → **Configurações da Central de Confiabilidade**
3. **Configurações de Macro**
4. Selecione **"Habilitar todas as macros"** (para desenvolvimento)

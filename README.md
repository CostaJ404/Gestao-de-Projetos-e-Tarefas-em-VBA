# 📊 Sistema de Gestão de Projetos e Tarefas - VBA Excel

![VBA](https://img.shields.io/badge/VBA-Excel-green)
![Status](https://img.shields.io/badge/Status-Completo-success)
![License](https://img.shields.io/badge/License-MIT-blue)

## 🎯 Sobre o Projeto

Sistema completo de **Gestão de Projetos e Tarefas** desenvolvido em VBA para Microsoft Excel, criado para demonstrar habilidades avançadas em automação de processos e desenvolvimento de soluções corporativas.

### ✨ Destaques

- 🎨 **Interface Gráfica Profissional** com UserForms intuitivos
- 💾 **CRUD Completo** para projetos e tarefas
- 📈 **Dashboard Interativo** com gráficos dinâmicos
- 📊 **Relatórios Automatizados** com exportação para PDF
- ✅ **Validações Robustas** e tratamento de erros
- 🎯 **Código Documentado** seguindo boas práticas

---

## 🚀 Funcionalidades

### Gestão de Projetos
- ✅ Cadastro completo de projetos com validação de dados
- ✅ Controle de datas, orçamentos e responsáveis
- ✅ Acompanhamento de status e progresso
- ✅ Edição e exclusão com confirmação
- ✅ Listagem e filtros personalizados

### Gestão de Tarefas
- ✅ Vínculo de tarefas a projetos específicos
- ✅ Controle de prioridades (Baixa, Média, Alta, Crítica)
- ✅ Acompanhamento de horas estimadas vs. reais
- ✅ Status detalhado (Pendente, Em Andamento, Completa)
- ✅ Cálculo automático de progresso do projeto

### Dashboard e Relatórios
- 📊 Gráfico de pizza com status dos projetos
- 📊 Gráfico de barras com prioridades das tarefas
- 📊 Análise de performance (horas, variações)
- 📊 Identificação de tarefas atrasadas
- 📊 Cronograma visual de projetos
- 📄 Exportação para PDF
- 📄 Relatórios por cliente

---

## 💻 Tecnologias Utilizadas

- **Microsoft Excel** (versão 2016 ou superior)
- **VBA (Visual Basic for Applications)**
- **UserForms** para interface gráfica
- **Charts API** para gráficos
- **File System Objects** para exportação

---

## 📁 Estrutura do Projeto

```
Sistema-Gestao-Projetos/
│
├── modPrincipal.bas          # Módulo principal do sistema
├── modCRUD.bas               # Operações de banco de dados
├── modRelatorios.bas         # Geração de relatórios e gráficos
├── modMenu.bas               # Menu e procedimentos auxiliares
│
├── frmProjeto.frm            # Formulário de projetos
├── frmTarefa.frm             # Formulário de tarefas
│
├── GUIA_IMPLEMENTACAO.md     # Guia completo passo a passo
└── README.md                 # Este arquivo
```

---

## 🎬 Demonstração

### Tela Inicial - Menu Principal
Interface limpa e intuitiva com acesso rápido a todas as funcionalidades.

### Formulário de Projetos
Cadastro completo com validação em tempo real e formatação automática.

### Dashboard Interativo
Visualização de indicadores-chave com gráficos atualizados automaticamente.

---

## 📋 Pré-requisitos

- Microsoft Excel 2016 ou superior
- Macros habilitadas
- Conhecimento básico em Excel

---

## 🔧 Instalação

### Método Rápido

1. **Download**: Baixe todos os arquivos do projeto
2. **Abrir Excel**: Crie um novo arquivo Excel (.xlsm)
3. **Importar Módulos**: 
   - Pressione `Alt + F11`
   - Arquivo → Importar
   - Selecione todos os arquivos `.bas`
4. **Criar UserForms**:
   - Inserir → UserForm
   - Configure os controles conforme instruções
   - Cole o código dos arquivos `.frm`
5. **Inicializar**: Execute a macro `InicializarSistema`

### Guia Detalhado

Para instruções completas passo a passo, consulte o arquivo **[GUIA_IMPLEMENTACAO.md](GUIA_IMPLEMENTACAO.md)**

---

## 🎯 Como Usar

### 1. Inicialização
```vba
' Execute uma única vez ao configurar o sistema
InicializarSistema
```

### 2. Gerenciar Projetos
```vba
' Abrir formulário de projetos
AbrirFormularioProjetos
```

### 3. Gerenciar Tarefas
```vba
' Abrir formulário de tarefas
AbrirFormularioTarefas
```

### 4. Gerar Relatórios
```vba
' Atualizar dashboard e criar gráficos
GerarRelatorioCompleto
```

### 5. Exportar PDF
```vba
' Exportar dashboard para PDF
ExportarDashboardPDF
```

---

## 📊 Capturas de Tela

### Dashboard
- Indicadores gerais do sistema
- Gráficos de status e prioridades
- Análise de performance

### Formulários
- Interface limpa e profissional
- Validações em tempo real
- Feedback visual para o usuário

---

## 🎓 Conceitos Demonstrados

### Programação VBA
- ✅ Módulos e procedimentos
- ✅ UserForms e controles
- ✅ Eventos e callbacks
- ✅ Collections e Arrays
- ✅ Loops e estruturas de controle
- ✅ Error handling robusto
- ✅ Funções personalizadas

### Excel Avançado
- ✅ Manipulação de ranges
- ✅ Formatação condicional programática
- ✅ Criação de gráficos dinâmicos
- ✅ Validação de dados
- ✅ Exportação para diferentes formatos
- ✅ WorksheetFunction

### Boas Práticas
- ✅ Código modular e reutilizável
- ✅ Nomenclatura clara e consistente
- ✅ Documentação inline
- ✅ Separação de responsabilidades
- ✅ Validação de entrada do usuário
- ✅ Tratamento adequado de erros

---

## 🔒 Validações Implementadas

- 📅 Validação de datas (formato e consistência)
- 💰 Validação de valores numéricos
- 📝 Validação de campos obrigatórios
- 🔢 Validação de progresso (0-100%)
- 🔗 Validação de integridade referencial (projetos-tarefas)
- ⚠️ Confirmação para ações destrutivas

---

## 📈 Indicadores e Métricas

### Indicadores Gerais
- Total de projetos cadastrados
- Projetos ativos
- Tarefas pendentes
- Taxa de conclusão

### Análise de Performance
- Total de horas estimadas
- Total de horas reais
- Variação de horas
- Percentual de variação
- Tarefas no prazo vs atrasadas

---

## 🛠️ Melhorias Futuras

- [ ] Autenticação de usuários
- [ ] Notificações por e-mail
- [ ] Integração com Outlook Calendar
- [ ] Gráficos de Gantt avançados
- [ ] Módulo de equipe e recursos
- [ ] Histórico de alterações
- [ ] Backup automático
- [ ] Importação/Exportação de dados

---

## 🐛 Solução de Problemas

### "Macro não encontrada"
**Solução**: Verifique se todos os módulos foram importados corretamente no VBA Editor.

### "UserForm não encontrado"
**Solução**: Confirme que os UserForms foram criados com os nomes corretos: `frmProjeto` e `frmTarefa`.

### Gráficos não aparecem
**Solução**: Execute `GerarRelatorioCompleto` para criar os gráficos.

### Erro ao salvar
**Solução**: Certifique-se de salvar o arquivo como `.xlsm` (Habilitado para Macros).

---

## 📝 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

---

## 👨‍💻 Autor

**[Seu Nome]**

- LinkedIn: [Seu LinkedIn]
- GitHub: [Seu GitHub]
- Email: [Seu Email]

---

## 🤝 Contribuições

Contribuições são sempre bem-vindas! Sinta-se à vontade para:

1. Fork o projeto
2. Criar uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abrir um Pull Request

---

## 📚 Recursos Adicionais

- [Guia Completo de Implementação](GUIA_IMPLEMENTACAO.md)
- [Documentação VBA Microsoft](https://docs.microsoft.com/pt-br/office/vba/api/overview/excel)
- [Boas Práticas VBA](https://www.excel-pratique.com/en/vba/best-practices.php)

---

## ⭐ Agradecimentos

- Comunidade VBA por todo o conhecimento compartilhado
- Stack Overflow pelas soluções e discussões
- Microsoft pela documentação detalhada

---

## 📞 Suporte

Se você tiver alguma dúvida ou sugestão, sinta-se à vontade para:

- Abrir uma [Issue](https://github.com/seuusuario/seuprojeto/issues)
- Enviar um e-mail
- Conectar-se no LinkedIn

---

**⚡ Desenvolvido com dedicação para demonstrar excelência em VBA e automação Excel**

---

### 🎯 Por que este projeto é ideal para portfólios?

1. **Demonstra Competência Técnica**: Mostra domínio de VBA e Excel avançado
2. **Resolve Problemas Reais**: Aplicável em diversos contextos corporativos
3. **Código Profissional**: Seguindo padrões e boas práticas da indústria
4. **Documentação Completa**: Facilitando compreensão e manutenção
5. **Interface Amigável**: Demonstrando preocupação com UX/UI
6. **Escalável**: Base sólida para expansões futuras

---

**Última atualização**: Fevereiro 2026

**Versão**: 1.0

**Status**: ✅ Projeto Completo e Funcional

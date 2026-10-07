# Plano de Implementação: Aba Dedicada de Estatísticas (Dashboard)

> **Agente:** `project-planner`  
> **Estratégia:** TDD & Arquitetura Limpa (`clean-code`, `python-patterns`)  
> **Status:** Pronto para Aprovação / Execução  

---

## 🎯 Objetivo

Criar uma aba própria e dedicada de **Estatísticas / Dashboard** como a primeira aba do arquivo Excel gerado pelo **Table Magic**, conferindo destaque executivo para métricas, volumes e distribuições (Tipo de Solicitação, Situação e Responsáveis), enquanto a aba de dados brutos é limpa com a remoção dos painéis que antes ficavam escondidos no rodapé após longas rolagens.

---

## 🏗️ Arquitetura e Decisões de Design

1. **Localização e Ordem das Abas**:
   - A nova aba se chamará **`Dashboard`** (ou **`Estatísticas`**) e será inserida na **posição 1** do workbook (`index=0`), abrindo como a visualização padrão.
   - Receberá o tema visual primário (Esmeralda Executivo ou tema correspondente).

2. **Componentes da Aba de Estatísticas**:
   - **Banner Hero (Linhas 1-2)**:
     - Título executivo (ex: `"Painel Executivo & Estatísticas"`).
     - Subtítulo com data/hora de exportação e resumo de cobertura dos dados.
   - **Cartões de KPI de Impacto (Linhas 4-6)**:
     - *Total de Registros / Processos*
     - *Total de Categorias / Tipos*
     - *Situações / Status Rastreados*
     - *Período de Referência*
   - **Tabelas de Contabilização e Distribuição (Linhas 8+)**:
     - **Painel 1 - Distribuição por Tipo/Categoria**: Coluna de Tipo, Quantidade, % do Total e Barra de Dados visual (*DataBarRule*).
     - **Painel 2 - Distribuição por Situação/Status**: Posicionado lado a lado ou em grade harmônica com larguras de coluna dimensionadas sob medida (sem restrições da tabela de dados).
     - **Painel 3 - Distribuição por Responsável** (quando houver múltiplos arquivos de responsáveis correlacionados).

3. **Limpeza da Aba de Dados Brutos**:
   - Em [`backend/excel/builder.py`](file:///c:/Users/muril/Desktop/Projetos/table-magic/backend/excel/builder.py):
     - A função [`_write_model_sheet`](file:///c:/Users/muril/Desktop/Projetos/table-magic/backend/excel/builder.py#L475) encerra no `TOTAL GERAL`.
     - O bloco de rodapé ([`_write_dynamic_panel`](file:///c:/Users/muril/Desktop/Projetos/table-magic/backend/excel/builder.py#L713-L738)) é removido da aba de dados brutos e transferido para a nova função `_write_dashboard_sheet`.
   - A faixa de KPIs rápidos no topo da tabela de dados (linhas 3-4) é mantida para consulta contextual durante filtros, ou ajustada de forma elegante.

4. **Compatibilidade com Excel e Arrays Dinâmicos**:
   - Manter compatibilidade com [`backend/excel/dynamic_array.py`](file:///c:/Users/muril/Desktop/Projetos/table-magic/backend/excel/dynamic_array.py) para que qualquer fórmula de spill (`UNIQUE/FILTER`) funcione tanto em versões recentes do Office 365 quanto em visualizadores convencionais.

---

## 📋 Decomposição de Tarefas (Bite-Sized Tasks)

### Tarefa 1: Criação da Branch e Setup de Testes Automatizados
- **Arquivos:**
  - Git Branch: `feature/dashboard-tab`
  - Criar: `backend/tests/test_excel_builder.py`
- **Ações:**
  - Criar e mudar para a branch `feature/dashboard-tab`.
  - Criar teste automatizado com `unittest` verificando a geração de um workbook com dados simulados e validando:
    1. Existência da aba `Dashboard` como primeira aba (`wb.sheetnames[0] == "Dashboard"`).
    2. Existência dos KPIs e tabelas de estatísticas na nova aba.
    3. Que a aba de dados termina no `TOTAL GERAL` sem blocos residuais abaixo.
- **Critério de Aceite:**
  - O teste falha inicialmente (Fase Red do TDD).

---

### Tarefa 2: Implementação da Função `_write_dashboard_sheet`
- **Arquivos:**
  - Modificar: `backend/excel/builder.py`
- **Ações:**
  - Criar a função `_write_dashboard_sheet(ws, main_table, data_sheet_name, all_tables, tab_color)`:
    - Configurar Chrome da aba (gridlines ocultas, cor da aba, orientação landscape).
    - Desenhar o Banner executivo independente.
    - Montar os cartões de KPI consolidados.
    - Construir as tabelas de distribuição (Tipo e Situação) com larguras de colunas otimizadas para leitura (ex: nomes de tipos não truncados, colunas numéricas alinhadas à direita, percentuais formatados como `0.0%`).
    - Aplicar *Data Bars* na coluna de quantidade.
- **Critério de Aceite:**
  - A função desenha a estrutura completa da aba sem erros de openpyxl.

---

### Tarefa 3: Integração no `build_workbook` e Limpeza da Aba de Dados
- **Arquivos:**
  - Modificar: `backend/excel/builder.py:713-738` e `backend/excel/builder.py:825-867`
- **Ações:**
  - No `build_workbook`:
    - Identificar a tabela principal.
    - Criar a aba `Dashboard` como a primeira aba do workbook (`wb.create_sheet(title="Dashboard", index=0)`).
    - Chamar `_write_dashboard_sheet`.
  - No `_write_model_sheet`:
    - Desativar a escrita dos painéis no rodapé (`_write_dynamic_panel` após o `grand_total_row`).
- **Critério de Aceite:**
  - O teste automatizado passa (Fase Green do TDD).

---

### Tarefa 4: Validação dos Arrays Dinâmicos e Fórmulas
- **Arquivos:**
  - Verificar/Ajustar: `backend/excel/dynamic_array.py`
- **Ações:**
  - Garantir que a injeção de dynamic arrays processe todas as abas (incluindo `Dashboard`).
  - Validar que as fórmulas com referências entre abas (ou cálculos estáticos consolidados) não gerem erros `#NAME?` ou `#REF!` no Excel.
- **Critério de Aceite:**
  - Geração de binário `.xlsx` sem avisos de integridade XML.

---

### Tarefa 5: Teste E2E com PDF Real do Projeto
- **Arquivos:**
  - Executar: `modelo/Setor21.pdf`
- **Ações:**
  - Rodar script de conversão de ponta a ponta gerando o arquivo `output_test.xlsx`.
  - Inspecionar a estrutura das abas geradas e o layout da nova aba `Dashboard`.
- **Critério de Aceite:**
  - O arquivo gerado abre perfeitamente, com a primeira aba exibindo os dados de forma limpa e destacada.

---

### Tarefa 6: Revisão de Qualidade e Finalização
- **Ações:**
  - Revisar padrões de código segundo `@clean-code`.
  - Garantir retrocompatibilidade com a API FastAPI e com o Frontend React.
  - Commits organizados na branch `feature/dashboard-tab`.

---

## ❓ Portão Socrático & Alinhamento Final

Para garantir que o resultado final atenda exatamente às suas expectativas:
1. **Nome da Aba**: Prefere que seja **`Dashboard`** ou **`Estatísticas`**?
2. **Dados da Tabela Principal vs Múltiplos Arquivos**: Quando você carrega múltiplos PDFs (ex: Setor + colaboradores), a aba de estatísticas deve focar no arquivo principal (consolidado) ou apresentar também um comparativo entre as pessoas?
3. **Podemos prosseguir com a execução do plano seguindo o TDD?**

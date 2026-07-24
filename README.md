# Table Magic (v2.5 Studio)

Conversor profissional de tabelas em PDF para planilhas Excel (`.xlsx`), com unificação de múltiplos arquivos em abas personalizadas e painéis executivos.

A extração de tabelas é **100% local e segura** — sem dependência de APIs externas de IA. PDFs digitais (com texto selecionável) são lidos via `pdfplumber` / `PyMuPDF`, e PDFs escaneados passam por OCR automatizado com Tesseract.

---

## 🚀 Recursos Principais

- **Unificação Inteligente de PDFs em Abas**:
  - O 1º PDF é definido como a **Aba Principal (PDF Base)**.
  - Os PDFs seguintes se tornam **Abas Secundárias**, integrando dados e relacionamentos.
  - Nomes das abas editáveis diretamente na interface.
- **Interface Moderna & Flexível**:
  - Reordenação de PDFs via **Arraste & Solte (Drag & Drop)**.
  - Alternador de visualização entre **Modo Grade (Grid 2 colunas)** e **Modo Lista**.
  - Suporte completo a **Tema Escuro (Dark Mode)** e **Tema Claro (Light Mode)**.
- **Opções de Formatação Personalizadas**:
  - **Desabilitar Aba Principal**: Trata a primeira aba como uma aba secundária comum (desativando cruzamentos e opções dependentes).
  - **Ocultar Coluna "Responsável"**: Oculta no Excel a coluna gerada pelo cruzamento entre o processo principal e as abas secundárias.
  - **Ocultar Abas Secundárias**: Preserva os dados no arquivo final, mantendo apenas a aba principal visível por padrão.
- **Excel Executivo & Dinâmico**:
  - Fórmulas dinâmicas ativas (`SUBTOTAL`, `UNIQUE`, `FILTER`).
  - Faixa de KPIs e painéis de contabilização automática.
  - Estilização temática em verde esmeralda com congelamento de painéis (*freeze panes*) e autofiltro.

---

## 🛠️ Arquitetura & Tecnologias

- **Frontend**: React 19 + Vite + Tailwind CSS v4 + Motion (`src/`)
- **Backend**: FastAPI (Python 3.11+) (`backend/`) — extração de tabelas, OCR e geração do `.xlsx` via `openpyxl`

---

## 💻 Rodando Localmente

### Pré-requisitos
- **Node.js** (v18+)
- **Python** (v3.11+)
- **[Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)** instalado no sistema e adicionado ao `PATH`, com o pacote de idioma português (`por`) — *necessário apenas para processar PDFs escaneados/imagens*.

### Instalação

```bash
# 1. Instalar dependências do frontend
npm install

# 2. Criar e ativar o ambiente virtual Python
python -m venv .venv
.venv\Scripts\activate

# 3. Instalar dependências do backend Python
pip install -r requirements.txt
```

### Desenvolvimento
Inicia o Vite na porta `5173` e o servidor FastAPI na porta `8000` em paralelo:

```bash
npm run dev
```

### Produção

```bash
npm run build
npm run start
```
O servidor FastAPI servirá os estáticos do frontend e as rotas de API em `http://localhost:3000`.

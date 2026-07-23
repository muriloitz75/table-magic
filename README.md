# Table Magic

Conversor de tabelas em PDF para planilhas Excel, com junção de vários PDFs em abas.
A extração de tabelas é 100% local — sem chamadas a nenhuma IA/LLM. PDFs digitais
(com texto selecionável) são lidos via `pdfplumber`/`PyMuPDF`; PDFs escaneados
passam por OCR com Tesseract.

## Arquitetura

- **Frontend**: React + Vite + Tailwind (`src/`)
- **Backend**: FastAPI (`backend/`) — extração de tabelas e geração do `.xlsx`

## Rodando localmente

**Pré-requisitos:**
- Node.js
- Python 3.11+
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) instalado e no `PATH`, com o pacote de idioma português (`por`) — necessário apenas para processar PDFs **escaneados**. PDFs digitais funcionam sem ele.

**Instalação:**

```bash
npm install
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

(Usar um ambiente virtual evita conflitos com outros projetos Python na máquina.)

**Desenvolvimento** (sobe o Vite na porta 5173 e o FastAPI na porta 8000, com proxy de `/api`):

```bash
npm run dev
```

**Produção:**

```bash
npm run build
npm run start
```

O servidor sobe em `http://localhost:3000`, servindo o frontend buildado e a API.

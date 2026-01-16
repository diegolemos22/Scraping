# 🕷️ WebScraping – Robôs Selenium + Orquestrador (Federal, Estadual e Municipal)

Este repositório contém um conjunto de robôs de **web scraping** desenvolvidos em **Python + Selenium**, voltados para a coleta automática de **atos normativos, orientações e notícias** dos seguintes portais:

- **IOB Online – Municipal**
- **IOB Online – Estadual**
- **IOB Online – Federal**
- **Portal SPED**
- **Portal CTe / MDF-e**
- **Portal SVRS – BPE e MDF-e**
- **Outras fontes tributárias**

Toda a execução pode ser controlada pelo **Orquestrador Sequencial**, que processa os scripts em ordem, gera logs, checkpoints, deduplica bases e produz resumos da execução.

---

# 📁 Estrutura Geral do Projeto

```
WebScraping/
│
├── 0_ORCHESTRATOR.py           # Orquestrador geral dos robôs
├── orchestrator.json           # Manifesto com a lista dos robôs e configurações
│
├── 1_BPE_MFE_ORIENTACOES.py    # Scraper SVRS – BPE/MDF-e
├── 2_SPED_FIRE.py              # Scraper SPED (Destaques)
├── 3_CHECKPOINT_ORIENTACOES.py # Scraper IOB – Estadual
├── 4_IOB_ORIENTACOES.py        # Scraper IOB – Federal
├── 5_CTE_ORIENTACOES.py        # Scraper Portal CT-e
├── 6_IOB_MUNICIPAL.py          # Scraper IOB – Municipal
│
├── logs/                       # Logs gerais e logs por step
│   ├── steps/
│   ├── run_YYYYMMDD_HHMMSS.log
│   ├── summary_*.json
│   └── summary_*.csv
│
└── checkpoint.json             # Controle para retomada automática
```

---

# 🚀 Como Executar

## ✔️ 1. Criar ambiente virtual
```bash
python -m venv .venv
.venv\Scripts ctivate
```

## ✔️ 2. Instalar dependências
```bash
pip install -r requirements.txt
```

## ✔️ 3. Configurar o arquivo `.ENV`
```
USER_OR=ORIENTACAO
PWD_OR=*****
IOB_EMAIL=seu_email
IOB_SENHA=sua_senha
```

## ✔️ 4. Executar o orquestrador (modo completo)
```bash
python 0_ORCHESTRATOR.py --manifest orchestrator.json
```

Outros modos:
- `--auto-discover`
- `--resume`
- `--only`
- `--dry-run`
- `--headless-all`

---

# 🤖 Descrição dos Robôs
(Conteúdo reduzido para manter o arquivo leve — igual ao enviado no chat.)

---

# ✔️ Pré-requisitos
- Python 3.9+
- Firefox + GeckoDriver
- Acesso aos portais oficiais
- SMTP porta 25 (para envio de e-mail)

---

# 🔒 Segurança
- Não subir `.ENV` no GitHub
- Não compartilhar credenciais
- Manter `.gitignore` atualizado

---

# 📜 Licença
Projeto interno para automação de processos da área.

# 🕵️‍♂️ Web Scraping IOB Alertas

Este projeto automatiza a extração de **atos normativos** do portal **IOB Online**, consolidando os dados em planilhas Excel e enviando por e-mail após deduplicação.

---

## 📌 Funcionalidades
- Login automático no portal IOB.
- Navegação até **Meu Espaço → Meus Alertas**.
- Clique no sino do alerta alvo (por nome ou índice).
- Acesso aos detalhes do dia (data atual ou ajustada).
- Extração de itens:
  - **Parser específico para blocos municipais** (ISSQN - UF - Município).
  - **Fallback genérico** para tabelas, artigos e listas.
- Consolidação em Excel com layout padronizado:
  - Colunas: `Ato`, `Descrição`, `Esfera`, `UF`, `Municipio`, `Data de extração`, `Data de publicação`, `Fonte`, `StatusCarga`.
- Deduplicação avançada ignorando `Data de extração`.
- Envio automático por e-mail com anexo da base deduplicada.

---

## 🛠 Tecnologias Utilizadas
- **Python 3.9+**
- **Bibliotecas**:
  - `selenium` (automação web)
  - `pandas` (manipulação de dados)
  - `openpyxl` (Excel)
  - `smtplib` (envio de e-mail)
- **Firefox WebDriver** (Geckodriver)

---

## 📂 Estrutura do Código
- **Login e navegação**: funções `login_iob_simple`, `open_meu_espaco_and_click_meus_alertas`.
- **Extração**:
  - `extract_items_municipal_blocks` → parser municipal.
  - `extract_items_from_details_page` → fallback genérico.
- **Persistência**:
  - `save_to_excel_like_old` → salva base consolidada e backup.
  - `dedupe_base_excel` → remove duplicados.
- **Envio de e-mail**: `send_mail_with_attachment`.

---

## 🚀 Como Executar
1. **Clone o repositório**:
   ```bash
   git clone https://github.com/diegolemos22/Scraping.git
   cd Scraping
   ```
2. **Crie e ative um ambiente virtual**:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```
3. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```
4. **Configure credenciais**:
   - Crie um arquivo `.ENV` com:
     ```
     IOB_EMAIL=seu_email
     IOB_SENHA=sua_senha
     ```
5. **Execute o script**:
   ```bash
   python SCRAP_DOC_IOB_TAX.py
   ```

---

## ✅ Pré-requisitos
- Firefox instalado + Geckodriver compatível.
- Acesso ao portal IOB.
- Permissão para envio de e-mail via SMTP (porta 25).

---

## 📌 Observações
- O script utiliza **perfil real do Firefox** para evitar bloqueios.
- Caso ocorra CAPTCHA, será necessário intervenção manual.
- Layout final do Excel segue padrão definido internamente.

---

## 🔒 Segurança
- Nunca compartilhe credenciais ou tokens.
- Use `.gitignore` para ocultar arquivos sensíveis (`.ENV`, planilhas, etc.).

---

## 📄 Licença
Projeto interno para automação de processos. Uso restrito.

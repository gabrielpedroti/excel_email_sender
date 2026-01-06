# 📧 Excel Email Sender

Projeto desenvolvido a partir da necessidade de um cliente de **enviar e-mails informativos em massa** para uma grande lista de contatos armazenada em uma planilha Excel, de forma automatizada, segura e controlada.

O sistema realiza:
- Leitura da lista de contatos via Excel  
- Personalização do conteúdo do e-mail  
- Envio automático via SMTP  
- Geração de log detalhado com resumo da operação  

---

## 🛠️ Tecnologias utilizadas

- Python 3  
- Pandas  
- OpenPyXL  
- SMTP (envio de e-mails)  
- HTML para corpo do e-mail  
- python-dotenv para gerenciamento de variáveis sensíveis  

---

## 📂 Estrutura do projeto

```
excel_email_sender/
│
├── main.py
├── template.html
├── requirements.txt
├── .env.example
└── logs/
```

---

## ⚙️ Configuração do ambiente

Crie e ative o ambiente virtual:

```bash
python -m venv venv
```

Windows:
```bash
venv\Scripts\activate
```

Linux / Mac:
```bash
source venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

---

## 🧪 Configuração do projeto

Crie o arquivo `.env` a partir do `.env.example` e preencha:

```env
EMAIL=seu_email@dominio.com
SENHA=sua_senha
SMTP_SERVER=mail.seudominio.com
SMTP_PORT=465
USE_SSL=true

ARQUIVO_CONTATOS=contatos.xlsx
ASSUNTO=Assunto do e-mail
INTERVALO_ENVIO=3
```

---

## ▶️ Executando

```bash
python main.py
```

Ao final da execução, será gerado um log com:
- Total de registros
- E-mails enviados
- Registros sem e-mail
- Erros
- Duração do processo
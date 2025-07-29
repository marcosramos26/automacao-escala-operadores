# Automação de Convocação de Escalas Operacionais

Automatiza o envio mensal de escalas personalizadas para operadores de loja via e-mail e WhatsApp, com confirmação de presença centralizada em Google Forms.

---

## Conteúdo

- [Funcionalidades](#funcionalidades)
- [Tecnologias](#tecnologias)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Uso](#uso)

---

## Funcionalidades

- Envio automático de e-mails via Outlook a partir de planilha Excel (`Automatização Escalas.xlsx`, aba `STATUS OPERADORES`).
- Delay de 4 s entre envios e pausa de 30 s a cada 100 e-mails para evitar bloqueios.
- Envio de mensagens via WhatsApp Web como canal de emergência, fechando cada aba após envio.
- Contagem regressiva de 10 s antes do disparo no WhatsApp.
- Link para Google Forms em todos os canais para confirmação (CPF + Sim/Não).
- Uso de PROCV / ÍNDICE+CORRESP no Google Planilhas para atualizar automaticamente o status de respostas.

---

## Tecnologias

- Python 3.x  
- Pandas  
- win32com (Outlook)  
- pywhatkit, pyautogui (WhatsApp Web)  
- openpyxl  

---

## Pré-requisitos

- Windows 10 ou 11 com Outlook instalado e configurado  
- Google Chrome com WhatsApp Web autenticado  
- Planilha local `Automatização Escalas.xlsx`, aba `STATUS OPERADORES` com colunas:
  1. LOJA  
  2. DATA  
  3. HORA ENTRADA  
  4. HORA SAÍDA  
  5. DIA DA SEMANA  
  6. CPF  
  7. NOME  
  8. EMAIL  
  9. TELEFONE  
  10. STATUS  

---

## Instalação

1. Clone o repositório:  
   ```bash
   git clone https://github.com/marcosramos26/automacao-escala-operadores.git
   cd automacao-escala-operadores
(Opcional) Crie e ative um ambiente virtual:

bash
Copiar
Editar
python -m venv venv
venv\Scripts\activate   # Windows
source venv/bin/activate # Linux/Mac
Instale as dependências:

bash
Copiar
Editar
pip install -r requirements.txt
Uso
Atualize sua planilha no Google Planilhas; baixe-a como Automatização Escalas.xlsx e coloque na raiz do projeto.

Envio de e-mail
Execute:

bash
Copiar
Editar
python envio_email.py
Aguarde o script rodar (~4 s entre cada e-mail, pausa a cada 100).

Envio de WhatsApp (em caso de não-resposta)
Execute:

bash
Copiar
Editar
python envio_whatsapp.py
Deixe o WhatsApp Web aberto e em foco.

Aguarde a contagem regressiva de 10 s e o envio acontecer, com cada aba sendo fechada.

Confirmação via Google Forms

Todas as comunicações incluem um link para o Forms.

As respostas (Sim/Não) aparecem automaticamente na planilha de respostas do Forms.

Use PROCV ou ÍNDICE+CORRESP no Google Planilhas para preencher a coluna STATUS.

Estrutura do Projeto
bash
Copiar
Editar
automacao-escala-operadores/
│
├── Automatização Escalas.xlsx  # Planilha baixada do Google Sheets
├── envio_email.py             # Script de envio automático de e-mails
├── envio_whatsapp.py          # Script de envio de mensagens no WhatsApp Web
├── requirements.txt           # Dependências Python
└── README.md                  # Documentação (este arquivo)
Autor
Marcos Ramos
GitHub | LinkedIn
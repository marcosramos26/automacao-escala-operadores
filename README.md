# Automação de Convocação de Escalas Operacionais

Automatiza o envio mensal de escalas personalizadas para operadores de loja via e-mail e WhatsApp, facilitando a comunicação e agilizando o processo de confirmação de presença.

---

## Conteúdo

- [Funcionalidades](#funcionalidades)
- [Tecnologias](#tecnologias)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Uso](#uso)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Configurações](#configurações)
- [Considerações](#considerações)
- [Melhorias Futuras](#melhorias-futuras)
- [Autor](#autor)

---

## Funcionalidades

- Envio automático de e-mails via Outlook com escala personalizada.
- Envio de mensagens via WhatsApp Web reforçando a convocação.
- Contagem regressiva para preparar a tela do WhatsApp.
- Validação básica de números de telefone.
- Tratamento de exceções para evitar paradas inesperadas.

---

## Tecnologias

- Python 3.x
- Pandas
- pywhatkit
- pyautogui
- pywin32 (win32com)
- openpyxl

---

## Pré-requisitos

- Windows 10 ou 11 (Outlook instalado e configurado)
- Conta Outlook configurada no Windows
- Google Chrome instalado e WhatsApp Web autenticado

---

## Instalação

1. Clone o repositório:
    ```bash
    git clone https://github.com/seuusuario/automacao-escala.git
    cd automacao-escala
    ```

2. Crie um ambiente virtual (opcional mas recomendado):
    ```bash
    python -m venv venv
    source venv/bin/activate  # Linux/Mac
    venv\Scripts\activate     # Windows
    ```

3. Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

---

## Uso

1. Prepare o arquivo `escalas.xlsx` na pasta raiz com a aba `BASE CONVOCACAO`, contendo as colunas:
   - NOME, LOJA, DATA, DIA DA SEMANA, HORA ENTRADA, HORA SAÍDA, EMAIL, TELEFONE

2. Execute o script principal:
    ```bash
    python main.py
    ```

3. Para o envio do WhatsApp:
    - Deixe o WhatsApp Web aberto no Chrome.
    - Deixe a aba do WhatsApp Web ativa e em foco (tela principal).
    - Aguarde a contagem regressiva antes de iniciar o envio das mensagens.

---

## Estrutura do Projeto


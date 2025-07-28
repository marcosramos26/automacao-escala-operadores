import pandas as pd
import time
import win32com.client as win32
import pywhatkit as kit
import pyautogui

# === Carregar e preparar a planilha ===
df = pd.read_excel('escalas.xlsx', sheet_name='BASE CONVOCACAO')
df.columns = df.columns.str.strip()

df['DATA'] = pd.to_datetime(df['DATA'])
df['HORA ENTRADA'] = pd.to_datetime(df['HORA ENTRADA'], format='%H:%M:%S').dt.strftime('%H:%M')
df['HORA SAÍDA'] = pd.to_datetime(df['HORA SAÍDA'], format='%H:%M:%S').dt.strftime('%H:%M')

# === ENVIO DE E-MAILS ===
print("📧 Iniciando envio de e-mails...")
outlook = win32.Dispatch('outlook.application')

for i, row in df.iterrows():
    nome = row['NOME']
    loja = row['LOJA']
    data = row['DATA'].strftime('%d/%m/%Y')
    dia_semana = row['DIA DA SEMANA']
    entrada = row['HORA ENTRADA']
    saida = row['HORA SAÍDA']
    email = str(row['EMAIL']).strip()

    corpo_email = f"""
Olá {nome},

Você está escalado para a loja {loja} no dia {data} ({dia_semana}).

🕒 Entrada: {entrada}  
🕔 Saída: {saida}

Você confirma presença?

Por favor, responda este e-mail com *SIM* ou *NÃO*.

Atenciosamente,  
Equipe de Escalas
"""

    try:
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = f"Escala para o dia {data} - Loja {loja}"
        mail.Body = corpo_email.strip()
        mail.Send()
        print(f"✅ E-mail enviado para: {nome} ({email})")
        time.sleep(10)  # delay para evitar bloqueios
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {email}: {e}")

print("\n📨 Todos os e-mails foram processados!\n")

# === AVISO E CONTAGEM REGRESSIVA PARA WHATSAPP ===
print("⚠️ ATENÇÃO: Coloque a janela do WhatsApp Web em FOCO TOTAL (tela principal visível)")
print("Iniciando em 10 segundos... Prepare a tela!")

for i in range(10, 0, -1):
    print(f"{i}...", end='', flush=True)
    time.sleep(1)

print("\nIniciando envio de mensagens no WhatsApp!\n")

# === ENVIO DE WHATSAPP ===
for i, row in df.iterrows():
    nome = row['NOME']
    loja = row['LOJA']
    data = row['DATA'].strftime('%d/%m/%Y')
    dia_semana = row['DIA DA SEMANA']
    entrada = row['HORA ENTRADA']
    saida = row['HORA SAÍDA']
    telefone = str(row['TELEFONE']).strip().replace('-', '').replace(' ', '').replace('(', '').replace(')', '')

    if telefone.isdigit() and len(telefone) in [10, 11]:
        if len(telefone) == 10:
            telefone = telefone[:2] + '9' + telefone[2:]
        numero = f'+55{telefone}'

        mensagem = f"""
Olá {nome}, tudo bem?

Você está escalado para a loja {loja} no dia {data} ({dia_semana}).

🕒 Entrada: {entrada}  
🕔 Saída: {saida}

Por favor, confirme sua presença respondendo o e-mail que enviamos.

📧 Se não encontrar o e-mail na caixa de entrada, verifique também sua caixa de spam.

Agradecemos!
"""

        try:
            kit.sendwhatmsg_instantly(numero, mensagem.strip(), wait_time=15, tab_close=False)
            time.sleep(7)
            pyautogui.press('enter')
            time.sleep(0.5)
            pyautogui.press('enter')
            time.sleep(10)
            print(f"✅ Mensagem enviada para: {nome}")
        except Exception as e:
            print(f"❌ Erro ao enviar para {numero}: {e}")
    else:
        print(f"❌ Número inválido para {nome}: {telefone}")

print("✅ Todas as mensagens do WhatsApp foram enviadas.")

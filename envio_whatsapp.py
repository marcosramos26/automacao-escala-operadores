import pandas as pd
import time
import pywhatkit as kit
import pyautogui

# === Carregar e preparar a planilha ===
df = pd.read_excel('Automatização Escalas.xlsx', sheet_name='STATUS OPERADORES')
df.columns = df.columns.str.strip()

df['DATA'] = pd.to_datetime(df['DATA'])
df['HORA ENTRADA'] = pd.to_datetime(df['HORA ENTRADA'], format='%H:%M:%S').dt.strftime('%H:%M')
df['HORA SAÍDA'] = pd.to_datetime(df['HORA SAÍDA'], format='%H:%M:%S').dt.strftime('%H:%M')

print("⚠️ ATENÇÃO: Coloque a janela do WhatsApp Web em FOCO TOTAL (tela principal visível)")
print("Iniciando em 10 segundos... Prepare a tela!")

for i in range(10, 0, -1):
    print(f"{i}...", end='', flush=True)
    time.sleep(1)

print("\nIniciando envio de mensagens no WhatsApp!\n")

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

Por favor, confirme sua presença preenchendo o formulário abaixo:

🔗 

📧 Se não encontrar o e-mail na caixa de entrada, verifique também sua caixa de spam.

Agradecemos!
"""

        try:
            kit.sendwhatmsg_instantly(numero, mensagem.strip(), wait_time=15, tab_close=False)
            time.sleep(7)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(7)  # espera para garantir que a mensagem foi enviada
            pyautogui.hotkey('ctrl', 'w')  # fecha a aba atual no Chrome
            time.sleep(3)  # pequena pausa antes do próximo envio
            print(f"✅ Mensagem enviada para: {nome}")
        except Exception as e:
            print(f"❌ Erro ao enviar para {numero}: {e}")
    else:
        print(f"❌ Número inválido para {nome}: {telefone}")

print("✅ Todas as mensagens do WhatsApp foram enviadas.")

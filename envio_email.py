import pandas as pd
import time
import win32com.client as win32

# === Carregar e preparar a planilha ===
df = pd.read_excel('Automatização Escalas.xlsx', sheet_name='STATUS OPERADORES')
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

👉 Para aceitar ou recusar sua convocação, acesse o formulário abaixo, preencha com seu CPF e selecione SIM ou NÃO:

🔗 Formulário: https://forms.gle/Uhgp98hTx8cGFd9DA

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

        # Delay entre e-mails para evitar bloqueios
        time.sleep(4)

        # Pausa de segurança a cada 100 e-mails enviados
        if (i + 1) % 100 == 0:
            print("⏸️ Pausa de segurança de 30 segundos...")
            time.sleep(30)

    except Exception as e:
        print(f"❌ Erro ao enviar e-mail para {email}: {e}")

print("\n📨 Todos os e-mails foram processados com sucesso!")

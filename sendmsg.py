from pywhatkit import sendwhatmsg,sendwhatmsg_instantly
import pandas as pd
from datetime import date


#Defina a hora e minuto de envio
hora = int(input("Hora de Envio: ").strip())
minuto = int(input("Minuto de envio: ").strip())

#Mes atual
mesAtual = date.today().month

#Usando dateTime para ler o dia e mes atual. O menos 1 (Para igualar com os numeros do BD)
diaAtual = (date.today().day) - 1

#Lendo o banco de dados (Planilha e ABAS com meses)

#leia = pd.read_excel("niver.xlsx", sheet_name=f"MES{mesAtual}")
leia = pd.read_excel("niver.xlsx", sheet_name=f"MES{mesAtual}")

#Utilizando iloc [linhaInicio:linhaFim, Coluna].values[Retorna apenas o texto]
nome = leia.iloc[diaAtual:diaAtual + 1, 2].values[0]
print(nome)

#Fone
fone = leia.iloc[diaAtual:diaAtual + 1, 3].values[0]
print(str(fone).replace(".0", ""))

#Saída da mensagem final
msg = (f"{nome}, Parabéns pelo seu dia. Muita saúde e paz sempre e que nessa nova fase Deus continue abençoando seu caminho. Um grande abraço e feliz aniversário (:")
#msg = (f"{nome}, isto é uma mensagem automática, não responda.")
sendwhatmsg(f"+{int(fone)}", f'{msg}', hora,minuto+1)

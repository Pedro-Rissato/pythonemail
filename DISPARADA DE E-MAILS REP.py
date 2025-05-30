import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

CR2 = input("Digite a sigla de 2 da unidade")
CR3 = input("Digite a sigla de 3 da unidade")
REPO = input("Digite o REP do relógio")
HORA = input("Digite Bom dia ou Boa tarde")

email.To = "suporte14@polimix.com.br"
email.Subject = f"Teste - NOVO REP {CR2}/{CR3}"
email.HTMLBODY = f"""
<p> {HORA}!</p>

<p>Cadastramos um relógio na unidade de {CR2}/{CR3}, segue o REP para finalizar o cadastro: {REPO} </p>

<p>Att, </p>
"""
email.Send()
print("E-mail enviado com sucesso")

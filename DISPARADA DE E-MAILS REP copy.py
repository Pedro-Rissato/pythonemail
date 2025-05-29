import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

CR2 = AB1 # type: ignore
CR3 = CD2 # type: ignore
REPO = FG3 # type: ignore

email.To = "suporte14@polimix.com.br"
email.Subject = "Teste - NOVO REP {CR2}/{CR3}"
email.HTMLBODY = f"""
<p> Boa tarde!</p>

<p>Cadastramos um relógio na unidade de {CR2}/{CR3}, segue o REP para finalizar o cadastro: {REPO} </p>

<p>Att, </p>
"""
email.Send()
print("E-mail enviado com sucesso")
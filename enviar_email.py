import win32com.client as win32

#Criar a integração com o outlook
outlook = win32.Dispatch(' outlook.application')

# Criar um e-mail
email = outlook.CreateItem(0)

# Configurar as informações do seu e-mail
email.To = "e-mail do destinatário"
email.Subject = "assunto do e-mail"
email.HTMLBody = """corpo do e-mail
Exemplo do corpo do email:

<p>Olá, tudo bem?</p>

<p>Segue o relatório de vendas dos produtos referente ao último mês:</p>
  <ul>
    <li>Produto A: R$ 5.000,00</li>
    <li>Produto B: R$ 3.200,00</li>
    <li>Produto C: R$ 4.500,00</li>
  </ul>
  <p>Atenciosamente,<br>Equipe de Vendas</p>

"""

# Se quiser colocar um anexo no e-mail
anexo = "exemplo: C://Users/Documentos/arquivo.xlsx"

email.Attachments.Add(anexo)

#Enviar email
email.Send()
print("E-mail Enviado com Sucesso!")
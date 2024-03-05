import streamlit as st
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#side bar
with open('./emails.xlsx', 'rb') as modelo:
  modelo_byte = modelo.read()

with st.sidebar:
  st.markdown('## Faça o download do modelo aqui :point_down:')
  st.download_button(label='Baixar Modelo', data=modelo_byte, file_name='Modelo.xlsx', mime='application/vnd.ms-excel')

#parte principal
st.markdown('# Envio de Emails em massa :email:')

email = st.text_input('Email', placeholder='Digite aqui seu email..')
senha = st.text_input('Senha', placeholder='Digite aqui sua senha..', type='password')
assunto = st.text_input('Assunto', placeholder='Digite aqui o assunto..')
messagem = st.text_area('Messagem', placeholder='Digite aqui sua mensagem..')

#st.subheader('Carregar lista de emails')
st.markdown('## Carregar Lista de Email :notebook:')
arquivo = st.file_uploader(
  'Suba o arquivo excel aqui. Caso não tenha um modelo, faça download na barra lateral',
  type=['xlsx']
)
#clintes = pd.read_excel(arquivo)

# Chegando arquivo
if arquivo:
  clientes = pd.read_excel(arquivo)
  st.write('Lista de email para envio:')
  st.dataframe(clientes)
else:
  st.error('Nenhum arquivo foi carregado')


# Função de envio
def sendmail():
  for index, cliente in clientes.iterrows():
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = email
    msg['To'] = cliente['email']
    msgs = f'Olá, {cliente['nome']}\n \n{messagem}'
    msg.attach(MIMEText(msgs, 'plain'))

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email, senha)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()

# quando o botão for clicado
if st.button('Enviar Emails'):
  try:
    sendmail()
    st.success('Emails enviados com sucesso!')
  except Exception as e:
    st.error(f'Erro ao enviar os email: {e}')
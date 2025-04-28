from imbox import Imbox
import smtplib
from email.message import EmailMessage
import pandas as pd


class Email(object):

    def __init__(self, emails):
        self.emails = emails

    def definir_dados_conexao(arquivo_conexao):
        df = pd.read_excel(arquivo_conexao)

        host = df.loc[0, 'host']
        remetente = df.loc[0, 'email']
        senha = df.loc[0, 'password']
        enviado_por = df.loc[0, 'enviado_por']
        setor = df.loc[0, 'setor']
        empresa = df.loc[0, 'empresa']

        return host, remetente, senha, enviado_por, setor, empresa

    def criar_conexao(host, remetente, senha):
        with Imbox(host, username=remetente, password=senha):
            mensagem = 'Conexão estabelecida com sucesso!'
        return mensagem

    def criar_mensagem(
            remetente, senha, recebedor, mes, local_arquivo, func, arquivo,
            tipo, enviado_por, setor, empresa, host):

        msg = EmailMessage()
        msg['Subject'] = f'{func} {tipo} {mes}'
        msg['From'] = remetente
        msg['To'] = recebedor
        mensagem = f'''
            Olá {func},

            Segue {tipo} referente a
            {mes}

            {enviado_por}
            {setor}
            {empresa}
        '''
        msg.set_content(mensagem)

        with open(local_arquivo, 'rb') as f:
            file_data = f.read()
            file_name = arquivo

        msg.add_attachment(
            file_data, maintype='application', subtype='pdf',
            filename=file_name)

        with smtplib.SMTP_SSL(host, 465) as smtp:
            smtp.login(remetente, senha)
            smtp.send_message(msg)

        mensagem = 'Email enviado com sucesso!'
        return mensagem

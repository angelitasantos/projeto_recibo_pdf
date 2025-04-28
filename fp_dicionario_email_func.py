import os
import pandas as pd
from fp_base import carregar_arquivos_de_configuracao


class DicionarioEmail():

    def __init__(self, email_path):
        self.email_path = email_path

    def listar_emails_excel(
            planilha_path, nome_aba, coluna_nome, coluna_email):

        try:
            df = pd.read_excel(planilha_path, sheet_name=nome_aba)
        except ValueError as e:
            raise ValueError(f'A aba não foi encontrada. Erro: {e}')

        if coluna_nome not in df.columns or coluna_email not in df.columns:
            msg1 = f'As colunas {coluna_nome} e/ou {coluna_email}'
            msg2 = f'não foram encontradas na aba {nome_aba}.'
            raise ValueError(f'{msg1} {msg2}')

        dicionario_emails = dict(zip(df[coluna_nome], df[coluna_email]))

        return dicionario_emails

    def gerar_dicionario_emails(planilha_path, empresa1, empresa2):
        diretorio_atual = os.path.dirname(os.path.abspath(__file__))
        caminho = diretorio_atual + '\\' + planilha_path

        coluna_nome = 'NOME'
        coluna_email = 'EMAIL'

        dicionario_emails_emp1 = DicionarioEmail.listar_emails_excel(
            caminho, empresa1, coluna_nome, coluna_email)
        dicionario_emails_emp2 = DicionarioEmail.listar_emails_excel(
            caminho, empresa2, coluna_nome, coluna_email)

        return dicionario_emails_emp1, dicionario_emails_emp2


arquivos = carregar_arquivos_de_configuracao('config.txt')
planilha_path = arquivos['Excel']
empresa1 = arquivos['Empresa1']
empresa2 = arquivos['Empresa2']
dict_emails_emp1, dict_emails_emp2 = DicionarioEmail.gerar_dicionario_emails(
    planilha_path, empresa1, empresa2)

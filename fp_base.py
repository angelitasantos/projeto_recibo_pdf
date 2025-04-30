import os
import re
import pandas as pd
from PyPDF2 import PdfFileReader, PdfFileWriter
from fp_enviar_email import Email


class PdfComNomes:

    def __init__(
            self, arquivo_pdf, arquivo_excel, arquivo_selecao,
            arquivo_conexao, nome_empresa1, nome_empresa2,
            pasta_saida_base=''):
        self.pdf_path = arquivo_pdf
        self.excel_path = arquivo_excel
        self.selecao_path = arquivo_selecao
        self.conexao_path = arquivo_conexao
        self.empresa1 = nome_empresa1
        self.empresa2 = nome_empresa2
        self.pasta_saida_base = pasta_saida_base

        (self.empresa,
         self.tipo_recibo, self.enviar_email) = self.carregar_selecao()

        self.pasta_saida = os.path.join(self.pasta_saida_base, self.empresa)
        os.makedirs(self.pasta_saida, exist_ok=True)

        self.nomes_funcionarios = self.carregar_nomes()

    def carregar_selecao(self):
        with open(self.selecao_path, 'r') as file:
            conteudo = file.readlines()

        empresa = conteudo[0].split(':')[1].strip()
        tipo_recibo = conteudo[1].split(':')[1].strip()
        enviar_email = conteudo[2].split(':')[1].strip()

        return empresa, tipo_recibo, enviar_email

    def carregar_nomes(self):
        df = pd.read_excel(self.excel_path, sheet_name=self.empresa)
        nomes = df.iloc[:, 0].dropna().astype(str).tolist()
        return nomes

    def extrair_mes_ano(self, texto):
        padrao = r'(' \
          r'janeiro|fevereiro|março|abril|maio|junho|julho|agosto|' \
          r'setembro|outubro|novembro|dezembro) de \d{4}'
        resultado = re.search(padrao, texto, re.IGNORECASE)
        if resultado:
            mes_ano = resultado.group(0).strip()
            return mes_ano.title().replace(' ', '_')
        return 'Periodo_Desconhecido'

    def agrupar_por_funcionario(self):
        agrupados = []

        with open(self.pdf_path, 'rb') as f:
            pdf_reader = PdfFileReader(f)
            paginas_texto = []
            for i in range(pdf_reader.numPages):
                texto = pdf_reader.getPage(i).extractText().split('\n')
                paginas_texto.append(texto)

            i = 0
            while i < len(paginas_texto):
                texto_atual = ' '.join(paginas_texto[i])
                nome_encontrado = None

                for nome in self.nomes_funcionarios:
                    if nome in texto_atual:
                        nome_encontrado = nome
                        break

                if nome_encontrado:
                    paginas = [i]
                    if i + 1 < len(paginas_texto):
                        texto_proxima = ' '.join(paginas_texto[i + 1])
                        if nome_encontrado in texto_proxima:
                            paginas.append(i + 1)
                            i += 1

                    mes_ano = self.extrair_mes_ano(texto_atual)

                    agrupados.append({
                        'nome': nome_encontrado,
                        'paginas': paginas,
                        'mes_ano': mes_ano
                    })

                i += 1

        return agrupados

    def exportar_pdfs_por_funcionario(self):
        grupos = self.agrupar_por_funcionario()

        with open(self.pdf_path, 'rb') as f:
            pdf_reader = PdfFileReader(f)

            for grupo in grupos:
                nome = grupo['nome'].replace(' ', '_')
                periodo = grupo['mes_ano']

                tipo_recibo = self.tipo_recibo.replace(' ', '_')
                paginas = grupo['paginas']

                writer = PdfFileWriter()
                for i in paginas:
                    writer.addPage(pdf_reader.getPage(i))

                nome_arquivo = f'{nome} {tipo_recibo} {periodo}.pdf'

                nome_arquivo1 = nome_arquivo.replace('_', ' ')
                nome_arquivo = nome_arquivo1.replace(' De ', ' de ')

                caminho_saida = os.path.join(self.pasta_saida, nome_arquivo)

                with open(caminho_saida, 'wb') as output_file:
                    writer.write(output_file)

    def carregar_emails_do_excel(arquivo):
        df = pd.read_excel(arquivo)
        return {row['NOME']: row['EMAIL'] for _, row in df.iterrows()}

    @staticmethod
    def enviar_emails(diretorio_pdfs, email_flag, conexao_path, excel_path):
        if email_flag.lower() != 'sim':
            print('[ℹ] Envio de email desativado.')
            return

        print('[✉] Enviando emails...')

        emails_dict = PdfComNomes.carregar_emails_do_excel(excel_path)

        hst, eml, senha, env_por, setor, emp = Email.definir_dados_conexao(
            conexao_path)
        host = hst
        remetente = eml
        senha = senha
        enviado_por = env_por
        setor = setor
        empresa = emp
        Email.criar_conexao(host, remetente, senha)

        for nome, destinatario in emails_dict.items():
            nome_formatado = nome.replace(' ', ' ')
            for arquivo in os.listdir(diretorio_pdfs):
                if nome_formatado in arquivo:
                    caminho_pdf = os.path.join(diretorio_pdfs, arquivo)

                    partes = arquivo.split()
                    ano = partes[-1].replace('.pdf', '')
                    mes_ano = partes[-3]
                    mes = mes_ano + ' de ' + ano
                    p0 = partes[-4]
                    add = 'Adiantamento'
                    fpt = 'Pagamento'
                    fptc = 'Folha de Pagamento'
                    f13 = '13o Salario'
                    tipo = add if p0 == add else fptc if p0 == fpt else f13

                    try:
                        Email.criar_mensagem(
                            remetente, senha, destinatario,
                            mes, caminho_pdf, nome, arquivo, tipo,
                            enviado_por, setor, empresa, host)
                        print(f'[✔] Email enviado para {nome}')
                    except Exception as e:
                        print(f'[!] Erro ao enviar para {nome}: {e}')


def carregar_arquivos_de_configuracao(arquivo_config):
    arquivos = {}
    with open(arquivo_config, 'r') as f:
        for linha in f:
            chave, valor = linha.strip().split(':')
            arquivos[chave.strip()] = valor.strip()
    return arquivos


if __name__ == '__main__':
    arquivos = carregar_arquivos_de_configuracao('config.txt')

    pdf = PdfComNomes(
        arquivos['PDF'],
        arquivos['Excel'],
        arquivos['Selecao'],
        arquivos['Conexao'],
        arquivos['Empresa1'],
        arquivos['Empresa2']
    )
    pdf.exportar_pdfs_por_funcionario()

    PdfComNomes.enviar_emails(
        pdf.pasta_saida, pdf.enviar_email, pdf.conexao_path, pdf.excel_path)

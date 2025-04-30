import tkinter as tk
from tkinter import messagebox
import subprocess
import webbrowser
import os
from fp_base import carregar_arquivos_de_configuracao


class App:

    def __init__(self, root):
        self.root = root
        self.root.title('Escolha: Empresa, Tipo de Recibo e Enviar Email')
        self.centralizar_janela(self.root, 600, 200)

        self.empresa_var = tk.StringVar()
        self.recibo_var = tk.StringVar()
        self.enviar_email_var = tk.StringVar()

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_columnconfigure(2, weight=1)

        self.criar_interface()

    def centralizar_janela(self, janela, largura, altura):
        janela.update_idletasks()
        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()
        x = (largura_tela - largura) // 2
        y = (altura_tela - altura) // 2
        janela.geometry(f'{largura}x{altura}+{x}+{y}')

    def criar_interface(self):
        arquivos = carregar_arquivos_de_configuracao('config.txt')
        emp1 = arquivos['Empresa1']
        emp2 = arquivos['Empresa2']

        empresa_frame = tk.Frame(self.root, width=180)
        empresa_frame.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        tk.Label(empresa_frame, text='Escolha a EMPRESA:').grid(
            row=0, column=0, sticky='w')
        tk.Radiobutton(empresa_frame, text=emp1,
                       variable=self.empresa_var, value=emp1).grid(
                           row=1, column=0, sticky='w')
        tk.Radiobutton(empresa_frame, text=emp2,
                       variable=self.empresa_var, value=emp2).grid(
                           row=2, column=0, sticky='w')

        recibo_frame = tk.Frame(self.root, width=240)
        recibo_frame.grid(row=0, column=1, padx=10, pady=10, sticky='nsew')

        tk.Label(recibo_frame, text='Escolha o TIPO DE RECIBO:').grid(
            row=0, column=0, sticky='w')
        tipos = ['Folha de Pagamento', 'Adiantamento',
                 '13o Salario', 'Informe de Rendimentos (EM CONSTRUÇÃO)']
        for i, tipo in enumerate(tipos):
            tk.Radiobutton(recibo_frame, text=tipo, variable=self.recibo_var,
                           value=tipo).grid(row=i+1, column=0, sticky='w')

        email_frame = tk.Frame(self.root, width=180)
        email_frame.grid(row=0, column=2, padx=10, pady=10, sticky='nsew')

        tk.Label(email_frame, text='Enviar Email?').grid(
            row=0, column=0, sticky='w')
        tk.Radiobutton(email_frame, text='Sim',
                       variable=self.enviar_email_var, value='Sim').grid(
                           row=1, column=0, sticky='w')
        tk.Radiobutton(email_frame, text='Não',
                       variable=self.enviar_email_var, value='Não').grid(
                           row=2, column=0, sticky='w')

        botao_frame = tk.Frame(self.root)
        botao_frame.grid(row=1, column=0, columnspan=3, pady=20)

        tk.Button(botao_frame, text='Confirmar Seleção',
                  command=self.abrir_janela_confirmacao).grid(
                      row=0, column=2, padx=10)

        consultar_emp1 = 'Consultar ' + emp1
        consultar_emp2 = 'Consultar ' + emp2
        tk.Button(botao_frame, text=consultar_emp1,
                  command=lambda: self.iniciar_servidor('index1.html')).grid(
                      row=0, column=0, padx=10)
        tk.Button(botao_frame, text=consultar_emp2,
                  command=lambda: self.iniciar_servidor('index2.html')).grid(
                      row=0, column=1, padx=10)

    def abrir_janela_confirmacao(self):
        empresa = self.empresa_var.get()
        recibo = self.recibo_var.get()
        email = self.enviar_email_var.get()

        if not empresa or not recibo or not email:
            msg1 = 'Seleção Incompleta'
            msg2 = 'Por favor, selecione ambas as opções.'
            messagebox.showwarning(msg1, msg2)
            return

        confirm_win = tk.Toplevel(self.root)
        confirm_win.title('Confirmação')
        self.centralizar_janela(confirm_win, 300, 200)
        confirm_win.grab_set()

        msg1 = f'EMPRESA: {empresa}'
        msg2 = f'TIPO DE RECIBO: {recibo}'
        msg3 = f'ENVIAR EMAIL: {email}'
        texto = f'{msg1}\n{msg2}\n{msg3}'
        tk.Label(confirm_win, text='Confirme sua seleção:',
                 font=('Arial', 10, 'bold')).pack(pady=10)
        tk.Label(confirm_win, text=texto, justify='left').pack(pady=5)

        btn_frame = tk.Frame(confirm_win)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text='Confirmar', width=12,
                  command=lambda: self.confirmar(
                      confirm_win, empresa, recibo, email)).pack(
                          side='left', padx=10)
        tk.Button(btn_frame, text='Cancelar', width=12,
                  command=confirm_win.destroy).pack(side='right', padx=10)

    def confirmar(self, win, empresa, recibo, email):
        arquivos = carregar_arquivos_de_configuracao('config.txt')
        selecao_path = arquivos['Selecao']
        try:
            with open(selecao_path, 'w') as f:
                f.write(f'EMPRESA: {empresa}\n')
                f.write(f'TIPO DE RECIBO: {recibo}\n')
                f.write(f'ENVIAR EMAIL: {email}\n')

            subprocess.run(['python', 'fp_base.py'], check=True)

            messagebox.showinfo('Exportado', 'Script executado com sucesso!')
            win.destroy()
            self.root.destroy()
        except Exception as e:
            messagebox.showerror('Erro', f'Erro ao salvar: {e}')
            win.destroy()

    def iniciar_servidor(self, arquivo_html):
        caminho_projeto = os.path.dirname(os.path.abspath(__file__))
        os.chdir(caminho_projeto)

        subprocess.Popen(['python', '-m', 'http.server', '5500'])
        webbrowser.open(f'http://localhost:5500/{arquivo_html}')


if __name__ == '__main__':
    def on_closing():
        root.destroy()
        os._exit(0)

    root = tk.Tk()
    app = App(root)
    root.protocol('WM_DELETE_WINDOW', on_closing)
    root.mainloop()

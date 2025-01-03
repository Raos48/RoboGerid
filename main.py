import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading

class Aplicacao:

    def __init__(self):
        self.janela = tk.Tk()
        self.janela.title('Automação de Processos')
        self.janela.geometry('400x150')
        self.janela.resizable(False, False)

        # Variável para armazenar o caminho do arquivo
        self.caminho_arquivo = ''

        # Label para mostrar o arquivo selecionado
        self.label_arquivo = tk.Label(self.janela, text="Nenhum arquivo selecionado", font=('Arial', 10))
        self.label_arquivo.pack(pady=10)

        # Botão para selecionar o arquivo
        self.botao_selecionar = tk.Button(self.janela, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.botao_selecionar.pack()

        # Botão para iniciar a automação
        self.botao_iniciar = tk.Button(self.janela, text="Iniciar Automação", command=self.iniciar_automacao, state='disabled')
        self.botao_iniciar.pack(pady=10)

        # Barra de progresso
        self.barra_progresso = ttk.Progressbar(self.janela, orient='horizontal', length=300, mode='determinate')
        self.barra_progresso.pack(pady=10)

        # Inicializa a janela
        self.janela.mainloop()

    def selecionar_arquivo(self):
        self.caminho_arquivo = filedialog.askopenfilename(title="Selecionar arquivo", filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.caminho_arquivo:
            self.label_arquivo.config(text=f"Arquivo Selecionado: {self.caminho_arquivo}")
            self.botao_iniciar.config(state='normal')

    def iniciar_automacao(self):
        # Inicia a automação em uma thread separada
        t = threading.Thread(target=self.executar_automacao)
        t.start()

    def executar_automacao(self):
        # Aqui você chama a função do módulo bot_transbordo.py
        # e passa o caminho do arquivo selecionado
        try:
            from Pages.transbordo import executar_bot_transbordo
            executar_bot_transbordo(self.caminho_arquivo, self.barra_progresso)
            messagebox.showinfo("Sucesso", "Automação concluída com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")


if __name__ == "__main__":
    app = Aplicacao()

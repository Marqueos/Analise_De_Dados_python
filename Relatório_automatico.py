#### Código feito por Marcos Guilherme gevarowski Lampugnani ####
#### GITHUB: https://github.com/Marqueos ####
#### Linkedin: www.linkedin.com/in/marcos-guilherme-60895a22a ####


import os
import pdfrw
import datetime
import openpyxl
from tkinter import *
from tkinter import ttk, filedialog, Listbox, messagebox

def adicionar_pasta(variavel_pastas, listbox):
    pasta_selecionada = filedialog.askdirectory(mustexist=True)
    if pasta_selecionada:
        variavel_pastas.append(pasta_selecionada)
        listbox.insert("end", pasta_selecionada)
        print(f"Pastas selecionadas: {variavel_pastas}")

def excluir_pasta(variavel_pastas, listbox):
    selected_index = listbox.curselection()

    if selected_index:
        pasta_removida = variavel_pastas.pop(selected_index[0])
        listbox.delete(selected_index)
        print(f"Pasta removida: {pasta_removida}")
    else:
        messagebox.showinfo("Aviso", "Selecione uma pasta para excluir.")

def obter_data_criacao_pdf(arquivo):
    try:
        pdf = pdfrw.PdfReader(arquivo)
        info = pdf.Info
        if '/CreationDate' in info:
            creation_date = info['/CreationDate']
            creation_date = creation_date.replace('(', '').replace(')', '').replace('D:', '')
            creation_date = creation_date.replace(':', '')
            creation_date = creation_date.split('Z')[0].replace('-03\'00\'', '')
            while len(creation_date) < 14:
                creation_date += '0'
            dt = datetime.datetime.strptime(creation_date, '%Y%m%d%H%M%S')
            data_formatada = dt.strftime('%d/%m/%Y')
            return data_formatada
    except Exception as e:
        print(f"Erro ao processar o arquivo {arquivo}: {e}")
    return None

def cancelar_operacao():
    root.destroy()

def processar_pastas(pastas, palavras_chave):
    arquivos_pdf = []
    for pasta in pastas:
        for root, dirs, files in os.walk(pasta):
            for arquivo in files:
                if arquivo.endswith('.pdf') and any(keyword in arquivo for keyword in palavras_chave):
                    arquivos_pdf.append(os.path.join(root, arquivo))
    return arquivos_pdf

def processar_e_mostrar():
    palavras_chave_str = palavras_chave_entry.get()

    if palavras_chave_str:
        palavras_chave = [keyword.strip() for keyword in palavras_chave_str.split(',')]
    else:
        palavras_chave = []

    if not pastas_selecionadas:
        messagebox.showinfo("Aviso", "Adicione pelo menos uma pasta para processar.")
        return

    arquivos_pdf = processar_pastas(pastas_selecionadas, palavras_chave)
    criar_planilha(arquivos_pdf)

def criar_planilha(arquivos_pdf):
    if not arquivos_pdf:
        messagebox.showinfo("Aviso", "Nenhum arquivo PDF encontrado.")
        return

    planilha = openpyxl.Workbook()
    planilha_ativa = planilha.active
    planilha_ativa.append(['Arquivo', 'Data de Criação', 'Pasta'])

    soma_por_data_palavra_chave = {}

    for arquivo_pdf in arquivos_pdf:
        nome_pasta = os.path.basename(os.path.dirname(arquivo_pdf))
        data_criacao = obter_data_criacao_pdf(arquivo_pdf)
        if data_criacao:
            nome_arquivo = os.path.splitext(os.path.basename(arquivo_pdf))[0]
            planilha_ativa.append([nome_arquivo, data_criacao, nome_pasta])

            chave = (data_criacao, nome_pasta)
            soma_por_data_palavra_chave[chave] = soma_por_data_palavra_chave.get(chave, 0) + 1

    planilha_soma = openpyxl.Workbook()
    planilha_soma_ativa = planilha_soma.active
    planilha_soma_ativa.append(['Data', 'Pasta', 'Quantidade'])

    for chave, soma in soma_por_data_palavra_chave.items():
        planilha_soma_ativa.append([chave[0], chave[1], soma])

    if len(planilha_ativa['A']) > 1:
        planilha.save('RELATÓRIO_DADOS.xlsx')

    if len(planilha_soma_ativa['A']) > 1:
        planilha_soma.save('SOMA_POR_DATA_PALAVRA_CHAVE.xlsx')

    planilha.close()
    planilha_soma.close()

def fechar_janela():
    root.destroy()

# Interface gráfica
root = Tk()
root.title("Processador de PDFs")
root.geometry("350x400")

frm = ttk.Frame(root, padding=10)
frm.grid(row=0, column=0, sticky="nsew")

ttk.Label(frm, text="Pastas Selecionadas").grid(column=0, row=0)

pastas_selecionadas = []
listbox = Listbox(frm, selectmode="extended", height=6, width=55)
listbox.grid(column=0, row=1, pady=10)

ttk.Label(frm, text="Palavras-chave (separadas por vírgula)").grid(column=0, row=2)
palavras_chave_entry = ttk.Entry(frm)
palavras_chave_entry.grid(column=0, row=3, pady=10)

botao_selecionar = ttk.Button(frm, text="Adicionar Pasta", command=lambda: adicionar_pasta(pastas_selecionadas, listbox))
botao_selecionar.grid(column=0, row=4, pady=8)

botao_excluir = ttk.Button(frm, text="Excluir Pasta", command=lambda: excluir_pasta(pastas_selecionadas, listbox))
botao_excluir.grid(column=0, row=5, pady=5)

botao_processar = ttk.Button(frm, text="Processar Pastas", command=processar_e_mostrar)
botao_processar.grid(column=0, row=6, pady=5)

botao_fechar = ttk.Button(frm, text="Fechar janela", command=fechar_janela)
botao_fechar.grid(column=0, row=7, pady=2)

botao_cancelar = ttk.Button(frm, text="Cancelar operação", command=cancelar_operacao)
botao_cancelar.grid(column=0, row=8, pady=25)

frm.columnconfigure(0, weight=1)
frm.rowconfigure(0, weight=1)

largura_janela = root.winfo_reqwidth()
altura_janela = root.winfo_reqheight()
posicao_x = (root.winfo_screenwidth() - largura_janela) // 2
posicao_y = (root.winfo_screenheight() - altura_janela) // 2
root.geometry("+{}+{}".format(posicao_x, posicao_y))

root.mainloop()

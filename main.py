import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def formatar_numeros(valor):
    if isinstance(valor, str):
        # Remover separadores de milhar e trocar vírgula decimal por ponto
        valor = valor.replace('.', '').replace(',', '.')
        try:
            return float(valor)  # Converter para float
        except ValueError:
            return valor  # Se não for possível converter, retorna o valor original
    elif isinstance(valor, (int, float)):
        return valor  
    return valor

def converter_para_csv(caminho_excel, caminho_csv):
    try:
        df = pd.read_excel(caminho_excel)
        df = df.applymap(formatar_numeros)
        df.to_csv(caminho_csv, sep=';', index=False, header=False, float_format="%.2f")
        messagebox.showinfo("Sucesso", f"Arquivo CSV gerado em: {caminho_csv}")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Função para abrir o diálogo de seleção do arquivo Excel
def selecionar_arquivo():
    arquivo_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
    entrada_entry.delete(0, tk.END)
    entrada_entry.insert(0, arquivo_excel)

# Função para escolher onde salvar o CSV
def salvar_como():
    caminho_csv = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="complementar.csv", filetypes=[("Arquivo CSV", "*.csv")])
    saida_entry.delete(0, tk.END)
    saida_entry.insert(0, caminho_csv)

# Função principal para rodar a conversão ao clicar no botão
def executar_conversao():
    caminho_excel = entrada_entry.get()
    caminho_csv = saida_entry.get()
    if caminho_excel and caminho_csv:
        converter_para_csv(caminho_excel, caminho_csv)
    else:
        messagebox.showwarning("Atenção", "Por favor, selecione o arquivo Excel e o local para salvar o CSV.")

# Interface gráfica
app = tk.Tk()
app.title("Conversor Excel para CSV")

# Labels e campos de entrada
tk.Label(app, text="Arquivo Excel:").grid(row=0, column=0, padx=10, pady=10)
entrada_entry = tk.Entry(app, width=50)
entrada_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Selecionar", command=selecionar_arquivo).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Salvar como (CSV):").grid(row=1, column=0, padx=10, pady=10)
saida_entry = tk.Entry(app, width=50)
saida_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Escolher local", command=salvar_como).grid(row=1, column=2, padx=10, pady=10)

# Botão de conversão
tk.Button(app, text="Converter", command=executar_conversao).grid(row=2, column=1, padx=10, pady=20)

# Inicia a interface
app.mainloop()

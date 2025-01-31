import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# Cores do tema escuro
COR_DE_FUNDO = "#2d2d2d"
COR_TEXTO = "#ffffff"
COR_BOTAO_MAGENTA = "#8B008B"
COR_BOTAO_NORMAL = "#404040"

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cruzar Dados de Planilhas")
        self.configure(bg=COR_DE_FUNDO)
        self.geometry("600x300")
        
        self.caminho_arquivo1 = ""
        self.caminho_arquivo2 = ""
        
        self.criar_widgets()
        self.configurar_estilos()
        
    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TFrame', background=COR_DE_FUNDO)
        style.configure('TLabel', background=COR_DE_FUNDO, foreground=COR_TEXTO)
        style.configure('TButton', 
                        background=COR_BOTAO_NORMAL, 
                        foreground=COR_TEXTO,
                        borderwidth=0)
        style.map('TButton', 
                 background=[('active', COR_BOTAO_NORMAL), ('pressed', COR_BOTAO_NORMAL)])
        
        style.configure('Magenta.TButton', 
                        background=COR_BOTAO_MAGENTA, 
                        foreground=COR_TEXTO)
        style.map('Magenta.TButton', 
                 background=[('active', COR_BOTAO_MAGENTA), ('pressed', '#6A006A')])
        
        style.configure('TEntry', fieldbackground="#404040", foreground=COR_TEXTO)
        
    def criar_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(padx=20, pady=20, fill='both', expand=True)
        
        # Arquivo 1
        ttk.Label(main_frame, text="Arquivo 1:").grid(row=0, column=0, sticky='w', pady=5)
        self.entry_arquivo1 = ttk.Entry(main_frame, width=50)
        self.entry_arquivo1.grid(row=1, column=0, padx=(0, 10))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_arquivo1).grid(row=1, column=1)
        
        # Arquivo 2
        ttk.Label(main_frame, text="Arquivo 2:").grid(row=2, column=0, sticky='w', pady=5)
        self.entry_arquivo2 = ttk.Entry(main_frame, width=50)
        self.entry_arquivo2.grid(row=3, column=0, padx=(0, 10))
        ttk.Button(main_frame, text="Procurar", command=self.selecionar_arquivo2).grid(row=3, column=1)
        
        # Botão Iniciar
        self.btn_iniciar = ttk.Button(main_frame, 
                                    text="Iniciar Processamento", 
                                    style='Magenta.TButton',
                                    command=self.processar,
                                    state='disabled')
        self.btn_iniciar.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Status
        self.status_label = ttk.Label(main_frame, text="", foreground=COR_TEXTO)
        self.status_label.grid(row=5, column=0, columnspan=2)
        
    def verificar_arquivos(self):
        if self.entry_arquivo1.get() and self.entry_arquivo2.get():
            self.btn_iniciar['state'] = 'normal'
        else:
            self.btn_iniciar['state'] = 'disabled'
            
    def selecionar_arquivo1(self):
        self.caminho_arquivo1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.entry_arquivo1.delete(0, tk.END)
        self.entry_arquivo1.insert(0, self.caminho_arquivo1)
        self.verificar_arquivos()
        
    def selecionar_arquivo2(self):
        self.caminho_arquivo2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.entry_arquivo2.delete(0, tk.END)
        self.entry_arquivo2.insert(0, self.caminho_arquivo2)
        self.verificar_arquivos()
        
    def processar(self):
        try:
            self.status_label['text'] = "Processando..."
            self.update_idletasks()
            
            aba = "Recibo de Pagamento"
            caminho_saida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if not caminho_saida:
                return
                
            dados1 = processar_planilha(self.caminho_arquivo1, aba)
            dados2 = processar_planilha(self.caminho_arquivo2, aba)
            
            if isinstance(dados1, str) or isinstance(dados2, str):
                raise ValueError(dados1 if isinstance(dados1, str) else dados2)
                
            dados_cruzados = cruzar_dados(dados1, dados2)
            salvar_em_excel_com_estilo(dados_cruzados, caminho_saida)
            self.status_label['text'] = "Processamento concluído com sucesso!"
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_saida}")
            
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            self.status_label['text'] = "Erro durante o processamento"
            
        finally:
            self.btn_iniciar['state'] = 'normal'

def processar_planilha(caminho_arquivo, aba):
    """
    Processa uma planilha Excel para extrair dados de funcionários, descrições, vencimentos e descontos.
    """
    try:
        dados = pd.read_excel(caminho_arquivo, sheet_name=aba)
        dados.reset_index(drop=True, inplace=True)
        resultado = {}

        i = 0
        while i < len(dados):
            if str(dados.iloc[i, 3]).strip().lower() == "nome":
                nome_funcionario = str(dados.iloc[i + 1, 3]).strip()
                resultado[nome_funcionario] = []
                i += 3

                while i < len(dados):
                    descricao = str(dados.iloc[i, 3]).strip()
                    if descricao == "" or descricao.lower() == "nan":
                        break

                    vencimentos = dados.iloc[i, 25] if not pd.isna(dados.iloc[i, 25]) else 0
                    descontos = dados.iloc[i, 33] if not pd.isna(dados.iloc[i, 33]) else 0

                    resultado[nome_funcionario].append({
                        "Descrição": descricao,
                        "Vencimentos": vencimentos,
                        "Descontos": descontos
                    })
                    i += 1
            i += 1

        return resultado
    except Exception as e:
        return str(e)

def cruzar_dados(dados1, dados2):
    """
    Realiza o cruzamento de dados entre dois dicionários de funcionários.
    """
    try:
        linhas = []

        todos_os_funcionarios = set(dados1.keys()).union(set(dados2.keys()))

        for nome in sorted(todos_os_funcionarios):
            registros1 = dados1.get(nome, [])
            registros2 = dados2.get(nome, [])

            mapa_registros1 = {r["Descrição"]: r for r in registros1}
            mapa_registros2 = {r["Descrição"]: r for r in registros2}

            todas_as_descricoes = sorted(set(mapa_registros1.keys()).union(set(mapa_registros2.keys())))

            for descricao in todas_as_descricoes:
                vencimentos1 = mapa_registros1.get(descricao, {"Vencimentos": 0})["Vencimentos"]
                descontos1 = mapa_registros1.get(descricao, {"Descontos": 0})["Descontos"]

                vencimentos2 = mapa_registros2.get(descricao, {"Vencimentos": 0})["Vencimentos"]
                descontos2 = mapa_registros2.get(descricao, {"Descontos": 0})["Descontos"]

                diferenca_vencimentos = vencimentos1 - vencimentos2
                diferenca_descontos = descontos1 - descontos2

                linhas.append({
                    "Funcionário": nome,
                    "Descrição": descricao,
                    "Vencimentos_Arquivo1": vencimentos1,
                    "Descontos_Arquivo1": descontos1,
                    "Vencimentos_Arquivo2": vencimentos2,
                    "Descontos_Arquivo2": descontos2,
                    "Diferença_Vencimentos": diferenca_vencimentos,
                    "Diferença_Descontos": diferenca_descontos
                })

        return pd.DataFrame(linhas)
    except Exception as e:
        return str(e)
        
def aplicar_estilo(val):
    """
    Estilos aplicados individualmente às colunas:
    - Fundo amarelo para valores diferentes de zero.
    - Fonte vermelha + fundo amarelo para valores negativos.
    - Adiciona "+" para valores positivos na coluna de diferença.
    """
    if val < 0:
        return "background-color: yellow; color: red"  # Fundo amarelo e fonte vermelha
    elif val > 0:
        return "background-color: yellow"  # Fundo amarelo para valores positivos
    elif val != 0:
        return "background-color: yellow"  # Apenas fundo amarelo para valores diferentes de zero
    return ""  # Sem estilo para valores iguais a zero

def salvar_em_excel_com_estilo(dados_cruzados, caminho_saida):
    """
    Salva os dados cruzados em Excel com formatação avançada e ajuste automático de colunas
    """
    try:
        df = pd.DataFrame(dados_cruzados)
        styled_df = df.style.applymap(aplicar_estilo, subset=["Diferença_Vencimentos", "Diferença_Descontos"])

        with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
            styled_df.to_excel(writer, index=False, sheet_name='Comparativo')
            
            workbook = writer.book
            worksheet = writer.sheets['Comparativo']
            
            # ========== FORMATOS PERSONALIZADOS ==========
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'fg_color': '#404040',
                'font_color': 'white',
                'border': 1,
                'border_color': '#707070'
            })

            money_format = workbook.add_format({
                'num_format': 'R$ #,##0.00',
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter'
            })

            default_format = workbook.add_format({
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter'
            })

            difference_format = workbook.add_format({
                'num_format': 'R$ #,##0.00;[Red]-R$ #,##0.00',
                'bold': True,
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter'
            })

            # ========== AJUSTE PRECISO DAS COLUNAS ==========
            for col_num, column in enumerate(df.columns):
                # Calcula a largura ideal considerando conteúdo formatado
                max_len = 0
                header_width = len(column) * 1.3  # Fator de conversão para largura Excel
                
                # Verifica o maior conteúdo da coluna
                if df[column].dtype == 'object':
                    max_data_width = df[column].astype(str).apply(len).max() * 1.2
                else:
                    max_data_width = df[column].astype(str).apply(len).max() * 1.5
                
                # Define a largura final com margem de segurança
                column_width = max(header_width, max_data_width) + 2
                
                # Aplica formatação específica
                if 'Vencimentos' in column or 'Descontos' in column:
                    worksheet.set_column(col_num, col_num, column_width, money_format)
                elif 'Diferença' in column:
                    worksheet.set_column(col_num, col_num, column_width + 3, difference_format)
                else:
                    worksheet.set_column(col_num, col_num, column_width, default_format)

            # ========== FORMATAÇÃO DO CABEÇALHO ==========
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # ========== RECURSOS AVANÇADOS ==========
            worksheet.set_default_row(28)  # Altura aumentada
            worksheet.autofilter(0, 0, 0, len(df.columns)-1)
            worksheet.freeze_panes(1, 0)
            worksheet.set_zoom(90)
            worksheet.set_tab_color('#FF69B4')  # Cor da aba em magenta

            # Adiciona legenda explicativa
            if len(df) > 0:
                worksheet.write(len(df)+2, 0, "Legenda:", header_format)
                worksheet.write(len(df)+3, 0, "Valores em amarelo", workbook.add_format({'bg_color': '#FFFF00'}))
                worksheet.write(len(df)+3, 1, "Diferenças significativas", default_format)
                worksheet.write(len(df)+4, 0, "Valores em vermelho", workbook.add_format({'font_color': '#FF0000'}))
                worksheet.write(len(df)+4, 1, "Diferenças negativas", default_format)

        print(f"Arquivo salvo com sucesso: {caminho_saida}")
        return True
    
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")
        return False

if __name__ == "__main__":
    app = Application()
    app.mainloop()
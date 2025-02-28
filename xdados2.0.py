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
                
                # Procurar pelo cabeçalho "Vencimentos" na linha logo abaixo do nome
                vencimentos_col = None
                descontos_col = None
                
                # Linha do cabeçalho está na posição i+2 (nome está em i+1)
                for col in range(dados.shape[1]):
                    if pd.notna(dados.iloc[i+2, col]) and str(dados.iloc[i+2, col]).strip().lower() == "vencimentos":
                        vencimentos_col = col
                    if pd.notna(dados.iloc[i+2, col]) and str(dados.iloc[i+2, col]).strip().lower() == "descontos":
                        descontos_col = col
                
                # Se não encontrou as colunas, use os valores padrão anteriores
                if vencimentos_col is None:
                    vencimentos_col = 25  # Coluna Z (padrão antigo)
                
                if descontos_col is None:
                    descontos_col = 33  # Coluna para descontos
                
                i += 3

                while i < len(dados):
                    descricao = str(dados.iloc[i, 3]).strip()
                    if descricao == "" or descricao.lower() == "nan":
                        break

                    vencimentos = dados.iloc[i, vencimentos_col] if not pd.isna(dados.iloc[i, vencimentos_col]) else 0
                    descontos = dados.iloc[i, descontos_col] if not pd.isna(dados.iloc[i, descontos_col]) else 0

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
    """
    try:
        val = float(val)
        if val < 0:
            return "background-color: yellow; color: red"
        elif val > 0:
            return "background-color: yellow"
        elif val != 0:
            return "background-color: yellow"
    except:
        return ""
    return ""

def salvar_em_excel_com_estilo(dados_cruzados, caminho_saida):
    """
    Salva os dados cruzados em Excel com formatação avançada e ajuste automático de colunas
    """
    try:
        df = pd.DataFrame(dados_cruzados)
        
        # Removemos a formatação manual do '+' e deixamos isso para o formato do Excel
        styled_df = df.style.map(aplicar_estilo, subset=["Diferença_Vencimentos", "Diferença_Descontos"])

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

            # Formato para números positivos (fundo amarelo)
            positive_format = workbook.add_format({
                'num_format': '+R$ #,##0.00',
                'bold': True,
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter',
                'bg_color': 'yellow',
                'font_color': 'black'
            })

            # Formato para números negativos (fundo vermelho, fonte branca)
            negative_format = workbook.add_format({
                'num_format': 'R$ #,##0.00',
                'bold': True,
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter',
                'bg_color': 'red',
                'font_color': 'white'
            })

            # Formato para zero
            zero_format = workbook.add_format({
                'num_format': 'R$ #,##0.00',
                'bold': True,
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter'
            })

            # Para células com fundo laranja
            orange_format = workbook.add_format({
                'bg_color': '#F79646',
                'font_color': 'white',
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter',
                'num_format': 'R$ +#,##0.00_);-R$ #,##0.00;R$ #,##0.00'
            })

            # Para células com fundo roxo
            purple_format = workbook.add_format({
                'bg_color': '#8064A2',
                'font_color': 'white',
                'border': 1,
                'border_color': '#D3D3D3',
                'valign': 'vcenter',
                'num_format': 'R$ +#,##0.00_);-R$ #,##0.00;R$ #,##0.00'
            })

            # ========== APLICAÇÃO DE CORES NAS LINHAS ==========
            for row_idx, row in df.iterrows():
                row_data = row.copy()
                
                # Colunas com índices das diferenças
                col_venc_diff = df.columns.get_loc("Diferença_Vencimentos")
                col_desc_diff = df.columns.get_loc("Diferença_Descontos")
                
                # Linha laranja: lançamento não identificado no arquivo 2
                if ((float(row['Vencimentos_Arquivo2']) == 0 and float(row['Vencimentos_Arquivo1']) > 0) or 
                    (float(row['Descontos_Arquivo2']) == 0 and float(row['Descontos_Arquivo1']) > 0)):
                    for col_idx, col_name in enumerate(df.columns):
                        value = row_data[col_name]
                        worksheet.write(row_idx + 1, col_idx, value, orange_format)
                
                # Linha roxa: lançamento não identificado no arquivo 1
                elif ((float(row['Vencimentos_Arquivo1']) == 0 and float(row['Vencimentos_Arquivo2']) > 0) or 
                      (float(row['Descontos_Arquivo1']) == 0 and float(row['Descontos_Arquivo2']) > 0)):
                    for col_idx, col_name in enumerate(df.columns):
                        value = row_data[col_name]
                        worksheet.write(row_idx + 1, col_idx, value, purple_format)
                
                # Linhas normais
                else:
                    for col_idx, col_name in enumerate(df.columns):
                        value = row_data[col_name]
                        if col_idx in [col_venc_diff, col_desc_diff]:
                            if float(value) > 0:
                                worksheet.write(row_idx + 1, col_idx, value, positive_format)
                            elif float(value) < 0:
                                worksheet.write(row_idx + 1, col_idx, value, negative_format)
                            else:
                                worksheet.write(row_idx + 1, col_idx, value, zero_format)
                        elif any(x in col_name for x in ['Vencimentos', 'Descontos']):
                            worksheet.write(row_idx + 1, col_idx, value, money_format)
                        else:
                            worksheet.write(row_idx + 1, col_idx, value, default_format)

            # ========== AJUSTE PRECISO DAS COLUNAS ==========
            for col_num, column in enumerate(df.columns):
                max_len = max(
                    len(str(column)),
                    df[column].astype(str).str.len().max()
                )
                adjusted_width = max_len * 1.2
                worksheet.set_column(col_num, col_num, adjusted_width)

            # ========== FORMATAÇÃO DO CABEÇALHO ==========
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # ========== RECURSOS AVANÇADOS ==========
            worksheet.set_default_row(28)
            worksheet.autofilter(0, 0, 0, len(df.columns)-1)
            worksheet.freeze_panes(1, 0)
            worksheet.set_zoom(90)
            worksheet.set_tab_color('#FF69B4')

            # Legenda atualizada
            if len(df) > 0:
                worksheet.write(len(df)+2, 0, "Legenda:", header_format)
                worksheet.write(len(df)+3, 0, "Valores em amarelo ", workbook.add_format({'bg_color': '#FFFF00',}))
                worksheet.write(len(df)+3, 1, "Diferenças a maior", default_format)
                worksheet.write(len(df)+4, 0, "Valores em vermelho", workbook.add_format({'bg_color': 'red', 'font_color': 'white'}))
                worksheet.write(len(df)+4, 1, "Diferenças a menor", default_format)
                worksheet.write(len(df)+5, 0, "Linha laranja", workbook.add_format({'bg_color': '#F79646', 'font_color': 'white'}))
                worksheet.write(len(df)+5, 1, "Lançamento não identificado na folha atual", default_format)
                worksheet.write(len(df)+6, 0, "Linha roxa", workbook.add_format({'bg_color': '#8064A2','font_color': 'white'}))
                worksheet.write(len(df)+6, 1, "Lançamento não identificado na folha anterior", default_format)

        print(f"Arquivo salvo com sucesso: {caminho_saida}")
        return True
    
    except Exception as e:
        print(f"Erro ao salvar o arquivo: {e}")
        return False
    
if __name__ == "__main__":
    app = Application()
    app.mainloop()
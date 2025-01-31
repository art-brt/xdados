import pandas as pd


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
    except FileNotFoundError:
        return "Erro: Arquivo não encontrado. Verifique o caminho do arquivo."
    except ValueError:
        return "Erro: Aba ou intervalo não encontrado no arquivo."
    except Exception as e:
        return f"Ocorreu um erro: {e}"


def cruzar_dados(dados1, dados2):
    """
    Realiza o cruzamento de dados entre dois dicionários de funcionários, incluindo todas as descrições de ambos os arquivos.
    """
    linhas = []

    # Itera sobre todos os funcionários presentes nos dois arquivos
    todos_os_funcionarios = set(dados1.keys()).union(set(dados2.keys()))

    for nome in sorted(todos_os_funcionarios):  # Ordena os funcionários alfabeticamente
        registros1 = dados1.get(nome, [])
        registros2 = dados2.get(nome, [])

        # Cria dicionários para mapear descrições em ambos os arquivos
        mapa_registros1 = {r["Descrição"]: r for r in registros1}
        mapa_registros2 = {r["Descrição"]: r for r in registros2}

        # Lista de todas as descrições presentes nos dois arquivos (ordenadas alfabeticamente)
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

    return linhas


def aplicar_estilo(val):
    """
    Aplica estilo de fundo amarelo para valores diferentes de zero.
    """
    if val != 0:
        return "background-color: yellow"
    return ""


def salvar_em_excel_com_estilo(dados_cruzados, caminho_saida):
    """
    Salva os dados cruzados em um novo arquivo Excel com formatação condicional.
    """
    try:
        df = pd.DataFrame(dados_cruzados)

        # Define o estilo condicional para as colunas de diferença
        styled_df = df.style.applymap(aplicar_estilo, subset=["Diferença_Vencimentos", "Diferença_Descontos"])

        # Salva o DataFrame estilizado em Excel
        styled_df.to_excel(caminho_saida, index=False, engine="openpyxl")

        print(f"Dados salvos com sucesso no arquivo: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao salvar os dados no arquivo Excel: {e}")


if __name__ == "__main__":
    # Caminhos dos arquivos Excel de entrada
    caminho_arquivo1 = "c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/Recibo de Pagamento dez.xlsx"
    caminho_arquivo2 = "c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/Recibo de Pagamento jan 426.xlsx"

    # Nome da aba
    aba = "Recibo de Pagamento"

    # Caminho do arquivo Excel de saída
    caminho_saida = "c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/dados_processados.xlsx"

    # Processa as duas planilhas
    dados_arquivo1 = processar_planilha(caminho_arquivo1, aba)
    dados_arquivo2 = processar_planilha(caminho_arquivo2, aba)

    # Realiza o cruzamento de dados e calcula as diferenças
    dados_cruzados = cruzar_dados(dados_arquivo1, dados_arquivo2)

    # Salva os dados cruzados em um novo arquivo Excel com formatação
    salvar_em_excel_com_estilo(dados_cruzados, caminho_saida)
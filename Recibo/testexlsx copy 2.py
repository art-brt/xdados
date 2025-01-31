import pandas as pd

def processar_planilha(caminho_arquivo, aba):
    """
    Processa uma planilha Excel para extrair dados de funcionários, descrições, vencimentos e descontos.
    
    :param caminho_arquivo: Caminho do arquivo Excel.
    :param aba: Nome ou índice da aba no arquivo Excel.
    :return: Um dicionário contendo os dados organizados.
    """
    try:
        # Lendo o arquivo Excel
        dados = pd.read_excel(caminho_arquivo, sheet_name=aba)

        # Garantindo que os índices comecem em 0
        dados.reset_index(drop=True, inplace=True)

        # Inicializa o dicionário final
        resultado = {}

        # Itera pelas linhas da planilha
        i = 0
        while i < len(dados):
            # Verifica na coluna D (indexada como 3 no Pandas) se há um campo "Nome"
            if str(dados.iloc[i, 3]).strip().lower() == "nome":
                # Pega o nome do funcionário (uma linha abaixo do campo "Nome")
                nome_funcionario = str(dados.iloc[i + 1, 3]).strip()

                # Cria o dicionário para este funcionário
                resultado[nome_funcionario] = []

                # Avança para a linha da descrição (duas linhas abaixo do campo "Nome")
                i += 2

                while i < len(dados):
                    # Pega o valor da descrição na coluna D
                    descricao = str(dados.iloc[i, 3]).strip()

                    # Se a descrição estiver em branco, interrompe o loop
                    if descricao == "" or descricao.lower() == "nan":
                        break

                    # Verifica os valores em "Vencimentos" (coluna Z, índice 25) ou "Descontos" (coluna AH, índice 33)
                    vencimentos = dados.iloc[i, 25] if not pd.isna(dados.iloc[i, 25]) else None
                    descontos = dados.iloc[i, 33] if not pd.isna(dados.iloc[i, 33]) else None

                    # Adiciona a descrição e o valor correspondente ao dicionário do funcionário
                    if vencimentos is not None:
                        resultado[nome_funcionario].append({"Descrição": descricao, "Valor": vencimentos, "Tipo": "Vencimentos"})
                    elif descontos is not None:
                        resultado[nome_funcionario].append({"Descrição": descricao, "Valor": descontos, "Tipo": "Descontos"})

                    # Vai para a próxima linha
                    i += 1

            # Vai para a próxima linha se "Nome" não estiver encontrado
            i += 1

        return resultado

    except FileNotFoundError:
        return "Erro: Arquivo não encontrado. Verifique o caminho do arquivo."
    except ValueError:
        return "Erro: Aba ou intervalo não encontrado no arquivo."
    except Exception as e:
        return f"Ocorreu um erro: {e}"


def salvar_em_excel(dados_processados, caminho_saida):
    """
    Salva os dados processados em um novo arquivo Excel.
    
    :param dados_processados: Dicionário com dados organizados.
    :param caminho_saida: Caminho para salvar o novo arquivo Excel.
    """
    try:
        # Cria uma lista para armazenar os dados em formato tabular
        linhas = []

        # Converte o dicionário para uma estrutura tabular
        for nome, registros in dados_processados.items():
            for registro in registros:
                linhas.append({
                    "Funcionário": nome,
                    "Descrição": registro["Descrição"],
                    "Tipo": registro["Tipo"],
                    "Valor": registro["Valor"]
                })

        # Cria um DataFrame com os dados
        df = pd.DataFrame(linhas)

        # Salva o DataFrame em Excel
        df.to_excel(caminho_saida, index=False)

        print(f"Dados salvos com sucesso no arquivo: {caminho_saida}")

    except Exception as e:
        print(f"Erro ao salvar os dados no arquivo Excel: {e}")


# Exemplo de uso
if __name__ == "__main__":
    # Caminho do arquivo Excel de entrada
    caminho_arquivo = "c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/Recibo-de-Pagamento_jan.xlsx"

    # Nome da aba
    aba = "Recibo de Pagamento"  # Substituir pelo nome ou índice da aba

    # Caminho do arquivo Excel de saída
    caminho_saida = "c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/dados_processados.xlsx"

    # Processa a planilha
    dados_processados = processar_planilha(caminho_arquivo, aba)

    # Exibe os dados processados de forma organizada
    if isinstance(dados_processados, dict):
        for nome, registros in dados_processados.items():
            print(f"\nFuncionário: {nome}")
            for registro in registros:
                print(f"  - Descrição: {registro['Descrição']}")
                print(f"    Tipo: {registro['Tipo']}")
                print(f"    Valor: {registro['Valor']}")

        # Salva os dados processados em um novo arquivo Excel
        salvar_em_excel(dados_processados, caminho_saida)
    else:
        print(dados_processados)
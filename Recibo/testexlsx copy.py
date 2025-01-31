import pandas as pd

def ler_celula(caminho_arquivo, aba, linha, coluna):
    """
    Lê uma célula específica de um arquivo Excel.
    
    :param caminho_arquivo: Caminho do arquivo .xlsx
    :param aba: Nome da aba ou índice da aba no arquivo Excel
    :param linha: Número da linha (base 1)
    :param coluna: Nome ou índice da coluna
    :return: Valor da célula
    """
    try:
        # Lê a aba especificada do Excel
        dados = pd.read_excel(caminho_arquivo, sheet_name=aba)

        # Como o Pandas trabalha com base 0, ajustamos a linha
        valor = dados.iloc[linha -1, coluna -1]  # Subtrai 1 de 'linha' para alinhar com o índice do Pandas

        return valor
    except FileNotFoundError:
        return "Erro: Arquivo não encontrado. Verifique o caminho do arquivo."
    except ValueError:
        return "Erro: Aba ou célula não encontrada no arquivo."
    except IndexError:
        return "Erro: Linha ou coluna fora do intervalo da tabela."
    except Exception as e:
        return f"Ocorreu um erro: {e}"

# Exemplo de uso da função
if __name__ == "__main__":
    # Caminho do arquivo Excel
    caminho_arquivo = 'c:/Users/Canella e Santos/Documents/arthur/xdados/Recibo/Recibo-de-Pagamento_jan.xlsx'

    # Nome da aba
    aba = 'Recibo de Pagamento'  # Substitua pelo nome ou índice da aba

    # Linha e coluna desejadas
    linha = 10  # Exemplo: segunda linha
    coluna = 4  # Exemplo: segunda coluna (índice baseado em 0 no Pandas)

    # Lê o valor da célula
    valor_celula = ler_celula(caminho_arquivo, aba, linha, coluna)
    print(f"Valor da célula na linha {linha}, coluna {coluna}: {valor_celula}")
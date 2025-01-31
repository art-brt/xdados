from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app)  # Permite requisições entre diferentes origens (React -> Flask)


@app.route("/upload", methods=["POST"])
def upload_files():
    if "file1" not in request.files or "file2" not in request.files:
        return jsonify({"error": "Por favor, envie ambos os arquivos!"}), 400

    file1 = request.files["file1"]
    file2 = request.files["file2"]

    try:
        # Lê os dois arquivos Excel
        dados1 = pd.read_excel(file1)
        dados2 = pd.read_excel(file2)

        # Processa os dois arquivos (substitua esta função pelo seu próprio processamento)
        resultado = processar_planilhas(dados1, dados2)

        # Retorna o resultado como JSON
        return jsonify({"message": "Arquivos processados com sucesso!", "data": resultado})

    except Exception as e:
        return jsonify({"error": f"Erro ao processar os arquivos: {str(e)}"}), 500


def processar_planilhas(dados1, dados2):
    """
    Função para cruzar e processar os dados de dois DataFrames.
    """
    resultado = []

    # Itera por todas as descrições presentes nos dois arquivos
    todas_as_descricoes = set(dados1["Descrição"]).union(set(dados2["Descrição"]))

    for descricao in todas_as_descricoes:
        vencimentos1 = dados1[dados1["Descrição"] == descricao]["Vencimentos"].sum() if descricao in dados1["Descrição"].values else 0
        descontos1 = dados1[dados1["Descrição"] == descricao]["Descontos"].sum() if descricao in dados1["Descrição"].values else 0

        vencimentos2 = dados2[dados2["Descrição"] == descricao]["Vencimentos"].sum() if descricao in dados2["Descrição"].values else 0
        descontos2 = dados2[dados2["Descrição"] == descricao]["Descontos"].sum() if descricao in dados2["Descrição"].values else 0

        resultado.append({
            "Descrição": descricao,
            "Vencimentos_Arquivo1": vencimentos1,
            "Descontos_Arquivo1": descontos1,
            "Vencimentos_Arquivo2": vencimentos2,
            "Descontos_Arquivo2": descontos2,
            "Diferença_Vencimentos": vencimentos1 - vencimentos2,
            "Diferença_Descontos": descontos1 - descontos2
        })

    return resultado


if __name__ == "__main__":
    app.run(debug=True)
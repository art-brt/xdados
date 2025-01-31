from flask import Flask, request, jsonify
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/process", methods=["POST"])
def process_excel_files():
    if 'baseFile' not in request.files or 'comparisonFile' not in request.files:
        return jsonify({"error": "Ambos os arquivos devem ser enviados"}), 400

    base_file = request.files['baseFile']
    comparison_file = request.files['comparisonFile']

    base_path = os.path.join(UPLOAD_FOLDER, base_file.filename)
    comparison_path = os.path.join(UPLOAD_FOLDER, comparison_file.filename)

    base_file.save(base_path)
    comparison_file.save(comparison_path)

    # Processamento de arquivos (substitua com suas funções reais)
    dados_base = pd.read_excel(base_path)
    dados_comparacao = pd.read_excel(comparison_path)

    resultado = {"base": len(dados_base), "comparacao": len(dados_comparacao)}

    return jsonify({"resultado": resultado})


if __name__ == "__main__":
    app.run(port=5000, debug=True)

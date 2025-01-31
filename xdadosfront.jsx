import React, { useState } from "react";

function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);

  const handleFile1Change = (event) => {
    setFile1(event.target.files[0]);
  };

  const handleFile2Change = (event) => {
    setFile2(event.target.files[0]);
  };

  const handleSubmit = async () => {
    if (!file1 || !file2) {
      alert("Por favor, insira ambos os arquivos!");
      return;
    }

    const formData = new FormData();
    formData.append("file1", file1);
    formData.append("file2", file2);

    setLoading(true);
    try {
      const response = await fetch("http://127.0.0.1:5000/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        alert(`Erro: ${errorData.error}`);
        setLoading(false);
        return;
      }

      const data = await response.json();
      setResult(data.data);
      setLoading(false);
    } catch (error) {
      alert("Erro ao enviar os arquivos!");
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-900 text-white flex flex-col items-center justify-center p-4">
      <h1 className="text-3xl font-bold mb-6">Processar Arquivos Excel</h1>

      <div className="space-y-4 w-full max-w-md">
        {/* Input para o Arquivo 1 */}
        <div>
          <label className="block text-lg font-medium mb-2" htmlFor="file1">
            Selecione o Arquivo 1:
          </label>
          <input
            type="file"
            id="file1"
            accept=".xlsx, .xls"
            onChange={handleFile1Change}
            className="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-gray-700 file:text-white hover:file:bg-gray-600"
          />
        </div>

        {/* Input para o Arquivo 2 */}
        <div>
          <label className="block text-lg font-medium mb-2" htmlFor="file2">
            Selecione o Arquivo 2:
          </label>
          <input
            type="file"
            id="file2"
            accept=".xlsx, .xls"
            onChange={handleFile2Change}
            className="block w-full text-sm text-gray-400 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-gray-700 file:text-white hover:file:bg-gray-600"
          />
        </div>

        {/* Botão de Enviar */}
        <button
          onClick={handleSubmit}
          className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-md"
          disabled={loading}
        >
          {loading ? "Processando..." : "Enviar"}
        </button>
      </div>

      {/* Exibição do Resultado */}
      {result && (
        <div className="mt-6 bg-gray-800 p-4 rounded-md text-sm max-w-2xl w-full">
          <h2 className="text-xl font-bold mb-4">Resultado:</h2>
          <pre className="whitespace-pre-wrap">{JSON.stringify(result, null, 2)}</pre>
        </div>
      )}
    </div>
  );
}

export default App;
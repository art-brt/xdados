import React, { useState } from 'react';
import { Button } from "@/components/ui/button";
import { Card, CardContent } from "@/components/ui/card";
import { FileInput } from "@/components/ui/file-input";
import { toast } from "@/components/ui/toast";
import { motion } from 'framer-motion';

const ExcelProcessor = () => {
  const [baseFile, setBaseFile] = useState(null);
  const [comparisonFile, setComparisonFile] = useState(null);

  const handleProcess = async () => {
    if (!baseFile || !comparisonFile) {
      toast.error("Selecione ambos os arquivos antes de iniciar o processamento.");
      return;
    }
  
    const formData = new FormData();
    formData.append("baseFile", baseFile);
    formData.append("comparisonFile", comparisonFile);
  
    try {
      const response = await fetch("http://localhost:5000/process", {
        method: "POST",
        body: formData
      });
  
      if (!response.ok) {
        throw new Error("Erro no processamento.");
      }
  
      const result = await response.json();
      toast.success(`Processamento concluído! Resultado: ${JSON.stringify(result.resultado)}`);
    } catch (error) {
      toast.error("Erro ao processar os arquivos.");
      console.error(error);
    }
  }

  return (
    <div className="min-h-screen bg-gray-900 text-white flex items-center justify-center p-6">
      <Card className="w-full max-w-lg bg-gray-800 shadow-lg rounded-2xl">
        <CardContent>
          <h2 className="text-2xl font-semibold text-center mb-4">Processador de Planilhas Excel</h2>
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium mb-1">Selecione a planilha base:</label>
              <FileInput onChange={e => setBaseFile(e.target.files[0])} />
              {baseFile && <p className="text-sm mt-2">Selecionado: {baseFile.name}</p>}
            </div>
            <div>
              <label className="block text-sm font-medium mb-1">Selecione a planilha de comparação:</label>
              <FileInput onChange={e => setComparisonFile(e.target.files[0])} />
              {comparisonFile && <p className="text-sm mt-2">Selecionado: {comparisonFile.name}</p>}
            </div>
            <div className="text-center mt-6">
              <Button
                disabled={!baseFile || !comparisonFile}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white p-3 rounded-xl"
                onClick={handleProcess}
              >
                Iniciar Processamento
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default ExcelProcessor;

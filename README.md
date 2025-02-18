# Comparador de Planilhas de Pagamento

Ferramenta GUI para comparar dados de pagamento entre duas planilhas Excel, destacando diferenças em vencimentos e descontos por funcionário.

## 📋 Funcionalidades
- **Interface Escura Moderna**: Design intuitivo com tema escuro.
- **Seleção de Arquivos**: Escolha duas planilhas Excel para comparação.
- **Processamento Automático**: Extrai e cruza dados da aba "Recibo de Pagamento".
- **Relatório Formatado**: Gera arquivo Excel com:
  - Cores para diferenças positivas/negativas.
  - Destaques para lançamentos ausentes.
  - Formatação monetária e ajuste automático de colunas.
  - Legenda detalhada.

## ⚙️ Requisitos
- Python 3.6+
- Bibliotecas: `pandas`, `openpyxl`, `xlrd`, `xlsxwriter`, `tkinter`

## 🛠 Instalação
```bash
pip install pandas openpyxl xlrd xlsxwriter

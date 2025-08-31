# 🖥️ Automação de Cadastro de Produtos

Este projeto automatiza o cadastro de produtos em um sistema web a partir de uma planilha Excel (`produtos_ficticios.xlsx`).  
Ele utiliza **pyautogui** para interagir com a interface gráfica e **openpyxl** para ler os dados da planilha.

⚠️ As coordenadas de clique (`pyautogui.click`) estão ajustadas para minha tela e podem precisar de ajuste em outros computadores.

---

## 🚀 Tecnologias usadas

- Python 3.11+
- [pyautogui]
- [openpyxl]
- [pyperclip]

## O que aprendi

- Usar a biblioteca PyAutoGUI para capturar posição do mouse.
- Copiar informações com pyperclip que copia até palavras com acentos.
- Percorrer cada linha de uma planilhar pegar os dados de cada célula.

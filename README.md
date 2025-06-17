# Sistema de Registro de Ponto

Este projeto é um sistema de registro de ponto personalizado que permite a marcação de horários de entrada, saída para almoço, retorno e saída do trabalho, com integração a planilhas Excel e exibição de informações climáticas via web scraping.

---

## 📦 Estrutura do Projeto

- `sistemaponto.py`: Interface gráfica e lógica principal de registro de ponto.
- `assets/`: Recursos visuais, como imagens e o layout `.ui` da interface gráfica.
- `Ponto Tecnologia.xlsx`: Planilha onde os dados de ponto são registrados.
- `backup/`: Diretório de backup automático da planilha.
- `import time.py`: Script comemorativo de teste.

---

## 🖥 Funcionalidades

- Interface gráfica feita com PyQt5.
- Registro de horários: entrada, saída para almoço, retorno e saída final.
- Integração com planilhas Excel para salvar dados.
- Cálculo automático de horas trabalhadas.
- Web scraping para exibir previsão do tempo atual de Juiz de Fora/MG.
- Backup automático da planilha ao registrar a saída.

---

## ▶️ Como Executar

1. Instale os requisitos:

```bash
pip install -r requirements.txt
```

2. Certifique-se de que o arquivo `Ponto Tecnologia.xlsx` e a interface `Interface_PONTO.ui` estão nos caminhos indicados no script (`D:/Executável/Ponto/`).

3. Execute o sistema:

```bash
python sistemaponto.py
```

---

## 🛠 Requisitos

- Python 3.x
- Biblioteca externas utilizadas:
  - `PyQt5`
  - `openpyxl`
  - `requests`
  - `beautifulsoup4`

---

## 📂 Organização de Pastas

```text
Sistema-Ponto/
├── sistemaponto.py
├── assets/
│   ├── Interface_PONTO.ui
│   ├── imagens, ícones, etc.
├── backup/
│   ├── Planilhas antigas de ponto
├── Ponto Tecnologia.xlsx
```

---

## 🧾 Licença

MIT

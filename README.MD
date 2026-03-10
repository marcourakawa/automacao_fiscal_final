# Automação Fiscal

Automação para preenchimento de planilhas e dados no sistema Protheus, com suporte a logs e validação de dados.  
O bot é feito em **Python** e pode ser convertido para `.exe` para rodar em computadores sem Python.

---

## 📝 Funcionalidades

- Leitura de planilhas Excel com dados fiscais.
- Formatação de campos para Protheus (datas, valores, códigos com zeros à esquerda).
- Preenchimento automático via **PyAutoGUI** e **Pyperclip**.
- Criação de logs detalhados em caso de erros.
- Console mostrando passo a passo do bot.
- Flexível: permite selecionar a planilha no momento da execução.

---

## 📂 Estrutura de Pastas

automacao_fiscal/
├─ planilha/ # Pasta contendo a planilha Excel
│ └─ planilha_uso_e_consumo.xlsx
├─ erro/ # Pasta onde os logs serão criados automaticamente
├─ _bot_diego_portavel.py # Script Python principal
└─ _bot_diego_portavel.exe # Executável gerado (opcional)

---

## ⚙️ Pré-requisitos (para rodar com Python)

- Python 3.10+  
- Bibliotecas:
  ```bash
  pip install pandas pyautogui pyperclip openpyxl
import os
import pandas as pd
import pyautogui
import pyperclip
import time
import re
from datetime import datetime
import sys
import tkinter as tk
from tkinter import filedialog

# ===== CONFIGURAÇÕES INICIAIS =====
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# ===== PASTA PARA LOG =====
pasta_base = os.path.dirname(os.path.abspath(sys.argv[0]))
pasta_log = os.path.join(pasta_base, "erro")
os.makedirs(pasta_log, exist_ok=True)
data_hora_execucao = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
arquivo_log = os.path.join(pasta_log, f"linhas_erradas_{data_hora_execucao}.txt")

# ===== SELEÇÃO DA PLANILHA =====
root = tk.Tk()
root.withdraw()  # esconde a janela principal
arquivo_excel = filedialog.askopenfilename(
    title="Selecione a planilha",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not arquivo_excel:
    print("Nenhuma planilha selecionada. Saindo...")
    exit(1)

# ===== LENDO A PLANILHA =====
try:
    df = pd.read_excel(arquivo_excel, dtype=str)
except Exception as e:
    print(f"Erro ao ler a planilha: {e}")
    exit(1)

# ===== Passo 2.2: Formatar data =====
if 'DT_EMISSAO' in df.columns:
    df['DT_EMISSAO'] = pd.to_datetime(
        df['DT_EMISSAO'], errors='coerce'
    ).dt.strftime('%d/%m/%Y')

# ===== Passo 2.3: Formatar VAL_UNIT como moeda brasileira =====
if 'VAL_UNIT' in df.columns:
    def formatar_val_unit(valor):
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        
        valor_str = str(valor).strip().replace(" ", "")
        
        try:
            # Caso venha no formato americano com ponto decimal (ex: 13216.58, 719.4)
            if "." in valor_str:
                numero_float = float(valor_str)
            else:
                # Caso seja inteiro grande (ex: 5214227 → 52.142,27)
                numero = int(valor_str)
                numero_float = numero / 100

            # Formata no padrão brasileiro
            return f"{numero_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        except Exception:
            # Se der erro, retorna valor original
            return valor_str

    df['VAL_UNIT'] = df['VAL_UNIT'].apply(formatar_val_unit)

# ===== Passo 2.4: Ajustar campos para Protheus (zeros à esquerda) =====
if 'NOTA_FISCAL' in df.columns:
    df['NOTA_FISCAL'] = df['NOTA_FISCAL'].str.zfill(9)

if 'COD_CLIEFOR' in df.columns:
    df['COD_CLIEFOR'] = df['COD_CLIEFOR'].str.zfill(8)

if 'COD_PRODUTO' in df.columns:
    df['COD_PRODUTO'] = df['COD_PRODUTO'].str.zfill(15)


# ===== Pega a primeira linha (teste) =====
if len(df) == 0:
    print("Erro: A planilha está vazia.")
    exit(1)

# ===== Ordem dos campos =====
CAMPOS_ORDEM = [
    "NOTA_FISCAL", "SERIE", "DT_EMISSAO", "COD_CLIEFOR",
    "ESPECIE", "COD_PRODUTO", "QTDE", "VAL_UNIT",
    "TES", "SOLIC_PAGTO", "NATUREZA", "CHAVE_NFE"
]

# ===== Função de preenchimento =====
def preencher_campos(linha, index):
    print("→ Iniciando em 2 segundos... (posicione a tela e NÃO mexa mais)")
    time.sleep(2)

    # ===== CONTEXTO DA LINHA =====
    chave_nfe = str(linha.get("CHAVE_NFE", "")).strip()
    if pd.isna(chave_nfe) or chave_nfe.lower() in ["nan", "none", "<na>", ""]:
        chave_nfe = ""

    print(f"→ Contexto | CHAVE_NFE={'SIM' if chave_nfe else 'NÃO'}")

    # =====================================================
    # FILIAL
    # =====================================================
    filial = str(linha.get("CNPJ_FILIAL", "")).strip()

    # FILIAL obrigatória
    if pd.isna(filial) or filial.lower() in ["nan", "none", "<na>", ""]:
        print("ERRO: FILIAL vazia. Linha será ignorada.")

        with open(arquivo_log, "a", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write(f"Data/Hora : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"Linha     : {index + 1}\n")
            f.write("Motivo    : FILIAL vazia\n")
            f.write("Dados da linha:\n")

            for coluna, valor in linha.items():
                f.write(f"  {coluna}: {valor}\n")

            f.write("\n")

        return

    # 🔒 TRAVA DE SEGURANÇA — padrão 00 0000
    if not re.match(r"^\d{2} \d{4}$", filial):
        print(
            f"ERRO: FILIAL fora do padrão ('{filial}'). "
            "Esperado: '00 0000'. Linha será ignorada."
        )

        with open(arquivo_log, "a", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write(f"Data/Hora : {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"!!!! LINHA !!!! = {index + 1}\n")
            f.write(f"Motivo    : FILIAL fora do padrão: '{filial}'\n")
            f.write("Dados da linha:\n")

            for coluna, valor in linha.items():
                f.write(f"  {coluna}: {valor}\n")

            f.write("\n")

        return


    pyautogui.press("i")
    time.sleep(4)

    pyautogui.press("tab")
    time.sleep(2)
    pyautogui.press("tab")
    time.sleep(2)

    pyperclip.copy(filial)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(1)

    for _ in range(4):
        pyautogui.press("tab")
        time.sleep(2)

    pyautogui.press("enter")
    time.sleep(5)

    # =====================================================
    # PREENCHIMENTO DOS CAMPOS
    # =====================================================


    for campo in CAMPOS_ORDEM:

        # Segurnaça para eivtar que o bot continue quando algo estiver errado.
        pyautogui.moveRel(0, 0)

        # 🔒 REGRA: só processa CHAVE_NFE se tiver valor
        if campo == "CHAVE_NFE" and not chave_nfe:
            print("→ CHAVE_NFE vazia, campo ignorado")
            continue

        valor = str(linha.get(campo, "")).strip()
        if pd.isna(valor) or valor.lower() in ['nan', 'na', 'none', '<na>', '']:
            valor = ""

        print(f"Preenchendo {campo:22}: {valor!r}")

        if campo == "NOTA_FISCAL":
            pyautogui.press("tab")
            time.sleep(3)
            pyautogui.press("tab")
            time.sleep(3)

            # Cola o valor da NOTA_FISCAL
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)

        elif campo == "SERIE":
            # Cola o valor da SERIE
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)

            # Se tiver menos que 3 caracteres, dar um tab extra
            if len(valor) < 3:
                pyautogui.press("tab")
                time.sleep(3)

        elif campo == "DT_EMISSAO":
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)

        elif campo == "COD_CLIEFOR":
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(2)
            pyautogui.press("tab")

        elif campo == "ESPECIE":
            # Cola o valor do campo ESPECIE
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)

            # Se valor tiver menos de 5 caracteres → 3 tabs, caso contrário → 2 tabs
            if len(valor) < 5:
                for _ in range(3):
                    pyautogui.press("tab")
                    time.sleep(2)
            else:
                for _ in range(2):
                    pyautogui.press("tab")
                    time.sleep(2)

            # Ações adicionais após tabs
            pyautogui.press("right")
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(3)


        elif campo == "COD_PRODUTO":
            # Cola o valor do campo COD_PRODUTO
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(6)

        elif campo == "QTDE":
            # Cola o valor do campo QTDE
            pyautogui.press("right")
            time.sleep(0.8)
            pyautogui.press("enter")
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(3)

        elif campo == "VAL_UNIT":
            for _ in range(3):
                pyautogui.press("right")
                time.sleep(2)

            pyautogui.press("enter")
            time.sleep(2)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(2)

            pyautogui.press("enter")
            time.sleep(2)


        elif campo == "TES":

            for _ in range(2):
                pyautogui.press("right")
                time.sleep(0.5)

            pyautogui.press("enter")
            time.sleep(2)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)

        elif campo == "SOLIC_PAGTO":
            
            for _ in range(22):
                pyautogui.press("right")
                time.sleep(0.8)

            pyautogui.press("enter")
            time.sleep(2)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            
            time.sleep(3)

        elif campo == "NATUREZA":
            print("→ Abrindo aba DUPLICATAS para preencher NATUREZA")

            # Clica na aba DUPLICATAS
            pyautogui.click(x=721, y=558)
            time.sleep(3)

            # Navega até o campo NATUREZA (2 tabs)
            for _ in range(2):
                pyautogui.press("tab")
                time.sleep(2)

            # Digita D20
            pyperclip.copy("D20")
            pyautogui.hotkey("ctrl", "v")
            time.sleep(2)

            pyautogui.press("tab")
            time.sleep(2)

            # Digita o valor
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(2)

            pyautogui.press("enter")
            time.sleep(3)

        elif campo == "CHAVE_NFE":
            print("→ Abrindo aba INFORMAÇÕES DANFE")
            time.sleep(3)

            # Clica na aba Informações DANFE
            pyautogui.click(x=1218, y=555)
            time.sleep(3)

            # Navega até o campo Chave NFE
            for _ in range(10):
                pyautogui.press("tab")
                time.sleep(0.8)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(4)
            pyautogui.press("enter")
            time.sleep(3)

            # Clicar no campo Tipo CT-e
            pyautogui.click(x=1044, y=667)
            time.sleep(3)
            pyautogui.press('n')
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(3)
            
    # Segurnaça para eivtar que o bot continue quando algo estiver errado.
    pyautogui.moveRel(0, 0)

    # Sair do loop salvar
    print("→ Clicando em salvar.")
    time.sleep(5)
    pyautogui.click(x=1309, y=146)
    time.sleep(20)

    # Bug protheus, cancelar a tela que abre sozinha.
    print("Fechando a tela incluir que foi aberta sozinha.")
    pyautogui.click(x=1227, y=152)
    time.sleep(15)

    # sair do loop cancelar
    # print("→ Clicando em cancelar")
    # pyautogui.click(x=1213, y=161)
    # time.sleep(10)

    print("\nPreenchimento da linha concluído.")

# ===== Executar =====
print("Automação pronta. Vai começar em 10 segundos...")
time.sleep(10)

try:
    for index, linha in df.iterrows():
        print(f"\n=== Processando linha {index + 1}/{len(df)} ===")
        preencher_campos(linha, index)
        time.sleep(5)  # pequeno respiro entre notas

except KeyboardInterrupt:
    print("\nAutomação interrompida pelo usuário (Ctrl+C)")
except Exception as e:
    print(f"\nERRO DURANTE A EXECUÇÃO:\n{e}")

print("\nAutomação finalizada para todas as linhas!")
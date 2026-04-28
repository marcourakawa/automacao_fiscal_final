import os # operacional system 
import pandas as pd # dataframe
import pyautogui # actions keyboard / mouse
import pyperclip # manipulate copy / paste
import time 
import re # regex
from datetime import datetime # time calc
import sys # python system      
import tkinter as tk # grafic interface
from tkinter import filedialog # open window to select an archive

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
            # Remove espaços e trata separadores
            if ',' in valor_str and '.' in valor_str:
                # Caso tenha ambos (ex: 1.234,56) → remove o ponto de milhar
                valor_limpo = valor_str.replace('.', '').replace(',', '.')
            elif ',' in valor_str:
                # Tem só vírgula → considera como decimal (padrão brasileiro)
                valor_limpo = valor_str.replace(',', '.')
            else:
                # Tem só ponto ou nenhum → trata normalmente
                valor_limpo = valor_str.replace(',', '.')
            
            numero_float = float(valor_limpo)
            
            # Formata como moeda brasileira
            return f"{numero_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        
        except Exception:
            return valor_str

    df['VAL_UNIT'] = df['VAL_UNIT'].apply(formatar_val_unit)   

# ===== Passo 2.4: Ajustar campos para Protheus (zeros à esquerda) =====
if 'NOTA_FISCAL' in df.columns:
    df['NOTA_FISCAL'] = df['NOTA_FISCAL'].str.zfill(9)

if 'COD_CLIEFOR' in df.columns:
    df['COD_CLIEFOR'] = df['COD_CLIEFOR'].str.zfill(8)

if 'COD_PRODUTO' in df.columns:
    df['COD_PRODUTO'] = df['COD_PRODUTO'].str.zfill(15)

if 'TES' in df.columns:
    df['TES'] = df['TES'].astype(str).str.zfill(3)

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
    time.sleep(2)

    # Clica na aba para inserir filial
    pyautogui.click(x=803, y=268)
    time.sleep(0.5)
    pyautogui.click(x=803, y=268)
    time.sleep(0.5)

    # Copiar e colocar a filial
    pyperclip.copy(filial)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    # Clicar em procurar
    pyautogui.click(x=976, y=261)
    time.sleep(1)

    # Clicar em ok
    pyautogui.click(x=344, y=580)
    time.sleep(4)


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
            # Clicar em NF
            pyautogui.click(x=984, y=211)
            time.sleep(1)
            pyautogui.press("backspace")

            # Cola o valor da NOTA_FISCAL
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

        elif campo == "SERIE":

            # Cola o valor da SERIE
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

            # Se tiver menos que 3 caracteres, dar um tab extra
            if len(valor) < 3:
                pyautogui.press("tab")
                time.sleep(1)

        elif campo == "DT_EMISSAO":
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

        elif campo == "COD_CLIEFOR":
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

        elif campo == "ESPECIE":
            pyautogui.press("tab")
            time.sleep(1)
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

        elif campo == "COD_PRODUTO":
            # Cola o valor do campo COD_PRODUTO
            pyautogui.click(x=28, y=305)
            time.sleep(1)
            pyautogui.press("right")
            time.sleep(0.8)
            pyautogui.press("enter")
            time.sleep(0.8)
            pyautogui.press("backspace")
            time.sleep(0.8)
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(5)

        elif campo == "QTDE":
            # Cola o valor do campo QTDE
            pyautogui.press("right")
            time.sleep(0.5)
            pyautogui.press("enter")
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)
            pyautogui.press("enter")
            time.sleep(1)

        elif campo == "VAL_UNIT":
            for _ in range(3):
                pyautogui.press("right")
                time.sleep(1)

            pyautogui.press("enter")
            time.sleep(1)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

            pyautogui.press("enter")
            time.sleep(1)


        elif campo == "TES":

            for _ in range(2):
                pyautogui.press("right")
                time.sleep(0.5)

            pyautogui.press("enter")
            time.sleep(1)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(2)

        elif campo == "SOLIC_PAGTO":
            
            for _ in range(22):
                pyautogui.press("right")
                time.sleep(0.5)

            pyautogui.press("enter")
            time.sleep(1)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

        elif campo == "NATUREZA":
            print("→ Abrindo aba DUPLICATAS para preencher NATUREZA")

            # Clica na aba DUPLICATAS
            pyautogui.click(x=711, y=554)
            time.sleep(1)

            # Navega até o campo NATUREZA (2 tabs)
            for _ in range(2):
                pyautogui.press("tab")
                time.sleep(1)

            # Digita D20
            pyperclip.copy("D20")
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

            pyautogui.press("tab")
            time.sleep(1)

            # Digita o valor
            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(1)

            pyautogui.press("enter")
            time.sleep(1)

        elif campo == "CHAVE_NFE":
            print("→ Abrindo aba INFORMAÇÕES DANFE")
            time.sleep(1)

            # Clica na aba Informações DANFE
            pyautogui.click(x=1196, y=557)
            time.sleep(1)

            # Navega até o campo Chave NFE
            pyautogui.click(x=752, y=591)
            time.sleep(1)

            pyautogui.press("backspace")
            time.sleep(1)

            pyperclip.copy(valor)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(3)
            pyautogui.press("enter")
            time.sleep(3)

            # Clicar no campo Tipo CT-e
            pyautogui.click(x=1043, y=666)
            time.sleep(2)
            pyautogui.press('n')
            time.sleep(2)
            pyautogui.press('enter')
            time.sleep(2)
            
    # Segurnaça para eivtar que o bot continue quando algo estiver errado.
    pyautogui.moveRel(0, 0)

    # TESTE CANCELAR
    # pyautogui.click(x=1232, y=150)
    # time.sleep(10)

    # Sair do loop salvar
    print("→ Clicando em salvar.")
    pyautogui.click(x=1312, y=149)
    time.sleep(10)
    
    if chave_nfe:
        print("CHAVE_NFE preenchida (CTE). Aguardando janela...")
        time.sleep(10)
        print("Tentando fechar janela popup do CTE")
        pyautogui.click(x=880, y=468)
        print("Janela fechada")
        time.sleep(3)

    # Bug protheus, cancelar a tela que abre sozinha.
    print("Fechando a tela incluir que foi aberta sozinha.")
    pyautogui.click(x=1232, y=150)
    time.sleep(5)

    print("\nPreenchimento da linha concluído.")

# ===== Executar =====
print("Automação pronta. Vai começar em 10 segundos...")
time.sleep(10)

try:
    for index, linha in df.iterrows():
        print(f"\n=== Processando linha {index + 1}/{len(df)} ===")
        preencher_campos(linha, index)
        time.sleep(3)  # pequeno respiro entre notas

except KeyboardInterrupt:
    print("\nAutomação interrompida pelo usuário (Ctrl+C)")
except Exception as e:
    print(f"\nERRO DURANTE A EXECUÇÃO:\n{e}")

print("\nAutomação finalizada para todas as linhas!")
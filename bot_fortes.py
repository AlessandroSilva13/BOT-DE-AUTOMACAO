import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyautogui
import time
import os
import getpass
from datetime import datetime


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5
USUARIO_ATUAL = getpass.getuser() # Ex: 'Joao', 'Maria'


def conectar_planilha():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
    client = gspread.authorize(creds)
    return client.open("Automação Fortes").get_worksheet(0)


def buscar_imagem(nome_arquivo, offset_x=0, tentativas=3):
    caminho = os.path.join('img', f'{nome_arquivo}.png')

    confianca = 0.8 
    
    for i in range(tentativas):
        try:
            posicao = pyautogui.locateCenterOnScreen(caminho, confidence=confianca)
            if posicao:
                pyautogui.click(posicao.x + offset_x, posicao.y)
                return True
        except:
            pass
        time.sleep(0.5)
    return False


def limpar_data(valor):
    return str(valor).replace('/', '').replace('.', '').replace('-', '').strip()
def ajustar_mes_ano(valor):
    limpo = str(valor).replace('/', '').replace('.', '').replace('-', '').strip()
    if len(limpo) == 8: return limpo[2:]
    return limpo
def digitar_enter(valor, espera=0.3):
    if valor == "" or valor is None:
        pyautogui.press('enter'); return
    pyautogui.write(str(valor))
    pyautogui.press('enter')
    time.sleep(espera)


print(f"\n=== ROBÔ FORTES | Usuário: {USUARIO_ATUAL} ===")
print("Este robô reserva os boletos na planilha para evitar conflitos.")
print("Mantenha o Fortes maximizado.")
time.sleep(5)

while True:
    try:
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Verificando fila...")
        planilha = conectar_planilha()
        lista_boletos = planilha.get_all_records()
        
        trabalhou = False

        for i, boleto in enumerate(lista_boletos, start=2):
            status = str(boleto.get('Status_Bot', ''))
            

            if status != "":
                continue


            print(f">>> Reservando linha {i} para {USUARIO_ATUAL}...")
            try:
                # Coluna 16 (P)
                planilha.update_cell(i, 16, f"Processando por {USUARIO_ATUAL}...")
                trabalhou = True
            except:
                print("Erro ao reservar linha. Talvez outro robô tenha pego. Pulando...")
                continue
            

            print(f"Processando Doc: {boleto.get('Documento')}")


            if not buscar_imagem('botao_incluir_principal', tentativas=2):

                pass
            time.sleep(2)


            buscar_imagem('campo.estabelecimento', offset_x=130)
            digitar_enter(boleto.get('Estabelecimento'))
            digitar_enter(boleto.get('Codigo'), espera=1.5)
            digitar_enter(boleto.get('Tipo_Doc'))
            digitar_enter(boleto.get('Documento'))
            digitar_enter(limpar_data(boleto.get('Data_Entrada')))
            
            val = str(boleto.get('Valor', '')).replace('.', ',')
            digitar_enter(val); pyautogui.press('enter'); time.sleep(0.5)
            
            digitar_enter(ajustar_mes_ano(boleto.get('Mes_Ano')))
            digitar_enter(limpar_data(boleto.get('Emissao_NF')))
            digitar_enter(str(boleto.get('Centro_Resultado')).zfill(7))
            digitar_enter(boleto.get('Despesa'))
            pyautogui.press('enter'); time.sleep(0.1)
            digitar_enter(boleto.get('Conta_Financeira'))
            pyautogui.press('enter'); time.sleep(0.1)
            pyautogui.write(str(boleto.get('Observacao', '')))
            pyautogui.press('enter'); time.sleep(0.5)


            if buscar_imagem('botao_novo'): time.sleep(1.5)
            else: pyautogui.press('f2'); time.sleep(1.5)

            pyautogui.write(limpar_data(boleto.get('Vencimento')))
            time.sleep(0.5)
            pyautogui.write(str(boleto.get('Codigo_Barras', '')).replace('.', '').replace(' ', '').replace('-', ''))
            time.sleep(1)


            if not buscar_imagem('botao_ok_f9'): pyautogui.press('f9')
            time.sleep(2)
            if not buscar_imagem('botao_ok_f9'): pyautogui.press('f9')
            time.sleep(4)


            planilha.update_cell(i, 16, f"Sucesso - {USUARIO_ATUAL}")
            print("Finalizado.\n")

        if not trabalhou:
            print("Nenhum boleto pendente. Aguardando 60s...")
            time.sleep(60)

    except Exception as e:
        print(f"Erro de conexão ou execução: {e}")
        time.sleep(30)
        import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyautogui
import time
import os
import getpass
from datetime import datetime


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5
USUARIO_ATUAL = getpass.getuser() # Ex: 'Joao', 'Maria'

def conectar_planilha():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
    client = gspread.authorize(creds)
    return client.open("Automação Fortes").get_worksheet(0)

def buscar_imagem(nome_arquivo, offset_x=0, tentativas=3):
    caminho = os.path.join('img', f'{nome_arquivo}.png')
    confianca = 0.8 
    
    for i in range(tentativas):
        try:
            posicao = pyautogui.locateCenterOnScreen(caminho, confidence=confianca)
            if posicao:
                pyautogui.click(posicao.x + offset_x, posicao.y)
                return True
        except:
            pass
        time.sleep(0.5)
    return False

def limpar_data(valor):
    return str(valor).replace('/', '').replace('.', '').replace('-', '').strip()
def ajustar_mes_ano(valor):
    limpo = str(valor).replace('/', '').replace('.', '').replace('-', '').strip()
    if len(limpo) == 8: return limpo[2:]
    return limpo
def digitar_enter(valor, espera=0.3):
    if valor == "" or valor is None:
        pyautogui.press('enter'); return
    pyautogui.write(str(valor))
    pyautogui.press('enter')
    time.sleep(espera)

print(f"\n=== ROBÔ FORTES | Usuário: {USUARIO_ATUAL} ===")
print("Este robô reserva os boletos na planilha para evitar conflitos.")
print("Mantenha o Fortes maximizado.")
time.sleep(5)

while True:
    try:
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Verificando fila...")
        planilha = conectar_planilha()
        lista_boletos = planilha.get_all_records()
        
        trabalhou = False

        for i, boleto in enumerate(lista_boletos, start=2):
            status = str(boleto.get('Status_Bot', ''))
            
            if status != "":
                continue


            print(f">>> Reservando linha {i} para {USUARIO_ATUAL}...")
            try:
                # Coluna 16 (P)
                planilha.update_cell(i, 16, f"Processando por {USUARIO_ATUAL}...")
                trabalhou = True
            except:
                print("Erro ao reservar linha. Talvez outro robô tenha pego. Pulando...")
                continue
            
            print(f"Processando Doc: {boleto.get('Documento')}")

            if not buscar_imagem('botao_incluir_principal', tentativas=2):
                pass
            time.sleep(2)


            buscar_imagem('campo.estabelecimento', offset_x=130)
            digitar_enter(boleto.get('Estabelecimento'))
            digitar_enter(boleto.get('Codigo'), espera=1.5)
            digitar_enter(boleto.get('Tipo_Doc'))
            digitar_enter(boleto.get('Documento'))
            data1 = str(boleto.get('Data_Entrada', '')).replace('.', '/')
            digitar_enter(data1);pyautogui.press('enter'); time.sleep(0.2)
            val = str(boleto.get('Valor', '')).replace('.', ',')
            digitar_enter(val); pyautogui.press('enter'); time.sleep(0.2)
            digitar_enter(ajustar_mes_ano(boleto.get('Mes_Ano')))
            digitar_enter(limpar_data(boleto.get('Emissao_NF')))
            digitar_enter(str(boleto.get('Centro_Resultado')).zfill(7))
            digitar_enter(boleto.get('Despesa'))
            pyautogui.press('enter'); time.sleep(0.1)
            digitar_enter(boleto.get('Conta_Financeira'))
            pyautogui.press('enter'); time.sleep(0.1)
            pyautogui.press('enter'); time.sleep(0.3)
            pyautogui.write(str(boleto.get('Observacao', '')))
            pyautogui.press('enter'); time.sleep(0.5)

            # 3. Vencimentos
            if buscar_imagem('botao_novo'): time.sleep(1.5)
            else: pyautogui.press('f2'); time.sleep(1.5)

            pyautogui.write(limpar_data(boleto.get('Vencimento')))
            time.sleep(0.5)
            pyautogui.write(str(boleto.get('Codigo_Barras', '')).replace('.', '').replace(' ', '').replace('-', ''))
            time.sleep(1)

            # 4. Finalizar
            if not buscar_imagem('botao_ok_f9'): pyautogui.press('f9')
            time.sleep(2)
            if not buscar_imagem('botao_ok_f9'): pyautogui.press('f9')
            time.sleep(4)

            # === SUCESSO ===
            planilha.update_cell(i, 16, f"Sucesso - {USUARIO_ATUAL}")
            print("Finalizado.\n")

        if not trabalhou:
            print("Nenhum boleto pendente. Aguardando 30s...")
            time.sleep(30)

    except Exception as e:
        print(f"Erro de conexão ou execução: {e}")
        time.sleep(30)
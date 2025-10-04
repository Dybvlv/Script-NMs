import win32com.client
import pandas as pd
import subprocess
import time
import os
from datetime import date
from datetime import datetime
import win32gui
import win32con
from openpyxl import load_workbook
import openpyxl 
import getpass
from pathlib import Path
import re

start_time = time.time()
# Detectando a chave do colaborador responsável pela execução do script
chave = os.getlogin()
# Caminho da pasta de salvamento dos arquivos
caminho_base = fr""


# Coletando data atual e formatando conforme é utilizado nos campos do SAP
hoje = date.today()
data_formatada = hoje.strftime("%d.%m.%Y")
print(f"Hoje é dia {data_formatada}")
variante = chave

def main():
    last_session = None
    session = None
    chave = os.getlogin()
    try:
        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        time.sleep(1)
        if SapGuiAuto:
            print("SAP GUI já está em execução.")
            application = SapGuiAuto.GetScriptingEngine
            if application.Children.Count > 0:
                last_connection = application.Children(application.Children.Count - 1)
                if last_connection.Children.Count > 0:
                    last_session = last_connection.Children(last_connection.Children.Count - 1)
                    print("Última janela detectada.")
                    last_session.findById("wnd[0]/tbar[0]/btn[419]").press()
                    time.sleep(5)
                    last_session.findById("wnd[0]").setFocus()
                    last_session.findById("wnd[0]").maximize()
                    print("A nova janela foi aberta com sucesso.") 
                    print("Janela maximizada.")
                    return last_session
                else:
                    print("Não há sessões abertas na última conexão.")
            else:
                print("Não havia conexões abertas no SAP GUI.")
        # Se não houver conexões, tenta abrir uma nova conexão
        connection = application.OpenConnection("02 PEP - SAP S/4HANA Produção (SAP SCRIPT)", True)
        last_session = connection.Children(0)
        last_session.findById("wnd[0]").setFocus()
        last_session.findById("wnd[0]").maximize()
        print("Conexão criada no SAP GUI.")
        return last_session
    except Exception as e:
        print(f"Erro detectado: {e}")
        print("Abrindo uma nova janela SAP...")
        sap_logon_path = r"C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
        subprocess.Popen(sap_logon_path)
        print("SAP Logon iniciado. Aguarde para conectar-se ao sistema SAP...")
        time.sleep(10)
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            application = SapGuiAuto.GetScriptingEngine
            print("Nova instância do SAP GUI detectada.")
            time.sleep(5)
            connection = application.OpenConnection("02 PEP - SAP S/4HANA Produção (SAP SCRIPT)", True)
            if connection:
                print("Conexão 02 PEP - SAP S/4HANA Produção (SAP SCRIPT) realizada com sucesso")
                last_session = connection.Children(0)
                last_session.findById("wnd[0]").setFocus()
                last_session.findById("wnd[0]").maximize()
            else:
                print("Erro ao conectar no 02 PEP - SAP S/4HANA Produção (SAP SCRIPT)")
        except Exception as e:
            print(f"Falha ao conectar ao SAP GUI após abertura: {e}")
    return last_session

# Função para fechar a janela do SAP
def close_sap_window(window_title):
    hwnd = win32gui.FindWindow(None, window_title)
    if hwnd:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(3)


def resolver_caminho_fotos_nms() -> Path:
    """Tenta localizar 'FOTOS NMs.xlsm' no OneDrive do usuário (Documents/Documentos ou busca recursiva)."""
    chave = getpass.getuser()
    base = Path(fr"")
    filename = "FOTOS NMs.xlsm"

    # 1) Testa as duas variações de pasta
    for docs in ("Documents", "Documentos"):
        p = base / docs / filename
        if p.exists():
            return p.resolve()

    # 2) Busca recursiva por nome de arquivo dentro do OneDrive
    if base.exists():
        print(f"Procurando por '{filename}' dentro de {base} ... (pode levar alguns segundos)")
        try:
            match = next(base.rglob(filename), None)
        except Exception:
            match = None
        if match:
            return match.resolve()

    # 3) Se não achar, deixa claro o problema
    raise FileNotFoundError(
        f"Não encontrei '{filename}' em:\n"
        f" - {base / 'Documents'}\n"
        f" - {base / 'Documentos'}\n"
        f"Dica: confira o nome exato da pasta e do arquivo no Explorer, "
        f"ou marque o arquivo no OneDrive como 'Sempre manter neste dispositivo'."
    )

# === Uso ===
excel_path = resolver_caminho_fotos_nms()
print("Planilha localizada em:", excel_path)

# Liste as abas para ter certeza do nome correto (muitas vezes .xlsm tem 'Planilha1' em PT-BR)
xls = pd.ExcelFile(excel_path, engine="openpyxl")
print("Abas encontradas:", xls.sheet_names)

# Leia a aba certa (ajuste 'Sheet1' se necessário)
Lista_NMS = pd.read_excel(excel_path, sheet_name="IH09", engine="openpyxl", dtype=str)


def ler_NMS_Materiais():
    
    Lista_NMS = pd.read_excel(fr"", sheet_name="")
    
    NMS_unicos = Lista_NMS['Lista_NMS'].unique().tolist()
    print(NMS_unicos)
    pd.Series(NMS_unicos).to_clipboard(index=False)
    pd.DataFrame(NMS_unicos).to_clipboard(index=False, header=False)

# ==== INÍCIO DO SCRIPT PRINCIPAL ====
last_session = main()
last_session.findById("wnd[0]").maximize()
last_session.findById("wnd[0]/tbar[0]/okcd").text = "/nih09"
last_session.findById("wnd[0]").sendVKey(0)
last_session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
last_session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 8
last_session.findById("wnd[0]/usr/btn%_MS_MATNR_%_APP_%-VALU_PUSH").press()
time.sleep(1)
ler_NMS_Materiais()
time.sleep(1)
last_session.findById("wnd[1]/tbar[0]/btn[16]").press()
last_session.findById("wnd[1]/tbar[0]/btn[24]").press()
last_session.findById("wnd[1]/tbar[0]/btn[8]").press()
last_session.findById("wnd[0]/tbar[1]/btn[8]").press()
grid = last_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
qtd_linhas_relatorio = grid.RowCount
print(f"A quantidade total é de {qtd_linhas_relatorio} NMS relacionados na lista disponibilizada.")

# --- Configurações ---
extensoes_permitidas = ('.jpg', '.jpeg')  # padronize
caminho_base = fr""  # ajuste para a sua base desejada

# --- Helpers ---
def safe_folder_name(s: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", str(s)).strip()

def safe_filename(name: str) -> str:
    n = re.sub(r'[\\/:*?"<>|]', "_", str(name)).strip()
    return n[:200]  # evita caminhos > 260 chars

def to_backslashes(path_str: str) -> str:
    return str(path_str).replace('/', '\\')

def wait_for_save_dialog(sess, timeout=5):
    """Retorna o índice da janela (1 ou 2) que contém DY_PATH."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        for idx in (2, 1):
            try:
                _ = sess.findById(f"wnd[{idx}]/usr/ctxtDY_PATH")
                return idx
            except:
                continue
        time.sleep(0.1)
    raise TimeoutError("Diálogo de salvar (ctxtDY_PATH) não apareceu a tempo.")

# --- Pasta base (sempre pasta, nunca arquivo) ---
if os.path.isfile(caminho_base):
    caminho_base = os.path.dirname(caminho_base)
caminho_base = os.path.join(caminho_base, "FOTOS_NMs")
os.makedirs(caminho_base, exist_ok=True)

# --- Leitura do GRID do IH09 ---
grid = last_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
qtd_linhas_relatorio = int(grid.RowCount)
print(f"Total de linhas no IH09: {qtd_linhas_relatorio}")

# --- Loop por cada linha do IH09 ---
for linha in range(qtd_linhas_relatorio):
    try:
        # Re-obter o grid a cada iteração (mais robusto após "voltar")
        grid = last_session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")

        numero_material = grid.GetCellValue(linha, "MATNR")
        descricao_material = grid.GetCellValue(linha, "MAKTG")

        # Normalização do NM para pasta
        numero_material_2 = str(numero_material).replace('.', '-')
        pasta_destino = os.path.join(caminho_base, safe_folder_name(numero_material_2))
        os.makedirs(pasta_destino, exist_ok=True)
        print(f"\n[Linha {linha}] NM: {numero_material} | Pasta: {os.path.normpath(pasta_destino)}")

        # --- Abrir detalhe do equipamento e GOS ---
        grid.currentCellRow = linha
        grid.doubleClickCurrentCell()
        last_session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
        last_session.findById("wnd[0]/shellcont/shell").pressButton("VIEW_ATTA")

        # --- Grid de anexos ---
        grid_anexos = last_session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
        # Aguarda RowCount
        for _ in range(50):
            try:
                qtd_linhas_anexos = int(grid_anexos.RowCount)
                break
            except:
                time.sleep(0.1)
        else:
            qtd_linhas_anexos = 0

        print(f"   Anexos encontrados: {qtd_linhas_anexos}")

        if qtd_linhas_anexos == 0:
            # Sem anexos — passa para o próximo NM (o finally fecha as janelas)
            continue

        # --- Loop dos anexos ---
        for linha_anexo in range(qtd_linhas_anexos):
            # 1) Nome do arquivo
            try:
                nome_arquivo = grid_anexos.GetCellValue(linha_anexo, 'BITM_FILENAME')
            except:
                time.sleep(0.3)
                nome_arquivo = grid_anexos.GetCellValue(linha_anexo, 'BITM_FILENAME')

            if not nome_arquivo:
                print(f"   [{linha_anexo}] Sem nome de arquivo. Pulando.")
                continue

            # 2) Filtrar extensão
            if not nome_arquivo.lower().endswith(extensoes_permitidas):
                print(f"   [{linha_anexo}] Ignorado (extensão não permitida): {nome_arquivo}")
                continue
            # 3) Caminho destino
            file_name = safe_filename(f"NM {numero_material_2} - Linha {linha_anexo} - {nome_arquivo}")
            destino_file = os.path.join(pasta_destino, file_name)

            if os.path.exists(destino_file):
                print(f"   [{linha_anexo}] Já existe: {os.path.normpath(destino_file)}")
                continue

            # 4) Exportar via menu de contexto
            try:
                grid_anexos.setCurrentCell(linha_anexo, 'BITM_DESCR')
                grid_anexos.selectedRows = str(linha_anexo)
                grid_anexos.contextMenu()
                grid_anexos.selectContextMenuItem("%ATTA_EXPORT")
            except Exception as e:
                print(f"   [{linha_anexo}] Falha ao acionar exportação: {e}")
                continue

            # 5) Preencher diálogo de salvar
            try:
                idx = wait_for_save_dialog(last_session, timeout=5)
                last_session.findById(f"wnd[{idx}]/usr/ctxtDY_PATH").text = to_backslashes(pasta_destino)
                last_session.findById(f"wnd[{idx}]/usr/ctxtDY_FILENAME").text = file_name
                last_session.findById(f"wnd[{idx}]/usr/ctxtDY_FILENAME").caretPosition = len(file_name)
                last_session.findById(f"wnd[{idx}]/tbar[0]/btn[0]").press()  # Salvar

                # Confirma overwrite, se aparecer
                try:
                    last_session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()  # "Sim"
                except:
                    pass
            except Exception as e:
                print(f"   [{linha_anexo}] Falha no diálogo de salvar: {e}")
                # Limpa diálogos residuais
                for i in (2, 1):
                    try:
                        last_session.findById(f"wnd[{i}]/tbar[0]/btn[12]").press()  # Cancel
                    except:
                        pass
                continue

            # 6) Verificar gravação
            if os.path.exists(destino_file):
                print(f"   [{linha_anexo}] OK: {os.path.normpath(destino_file)}")
            else:
                try:
                    msg = last_session.findById("wnd[0]/sbar").Text
                except:
                    msg = "-"
                print(f"   [{linha_anexo}] Atenção: arquivo não encontrado após salvar. Status: {msg}")

    except Exception as e:
        print(f"[Linha {linha}] Erro inesperado: {e}")

    finally:
        # --- Fecha anexos e volta para a lista do IH09 (por iteração) ---
        # fecha diálogos "Salvar como" se ainda abertos
        for i in (2, 1):
            try:
                last_session.findById(f"wnd[{i}]/tbar[0]/btn[12]").press()  # Cancel
            except:
                pass
        # fecha janela de anexos/volta
        for i in (1, 2):
            try:
                last_session.findById(f"wnd[{i}]/tbar[0]/btn[0]").press()  # OK/Voltar
            except:
                pass
        # voltar da tela do equipamento para a lista
        try:
            last_session.findById("wnd[0]/tbar[0]/btn[3]").press()  # Voltar
        except:
            pass
        time.sleep(0.2)  # pequeno respiro entre iterações

print("\n✅ Fim do processamento.")


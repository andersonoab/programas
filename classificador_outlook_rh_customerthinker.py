import os
import re
import json
import time
import threading
import traceback
from pathlib import Path
from datetime import datetime, date

import customtkinter as ctk
from tkinter import filedialog, messagebox

import pythoncom
import win32com.client as win32

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


APP_NAME = "Leitor Outlook RH - CustomerThinker"
APP_VERSION = "1.0.0"

BASE_DIR = Path(r"C:\_RPA\AppEmailRH")
SAIDA_DIR = BASE_DIR / "saida"
LOGS_DIR = BASE_DIR / "logs"
CONFIG_DIR = BASE_DIR / "config"
CONFIG_FILE = CONFIG_DIR / "config_email_rh.json"

DEFAULT_EXCEL_NAME = "base_atendimento_gente_gestao.xlsx"

CORES = {
    "navy": "#121C4E",
    "azul": "#0083CA",
    "azul_apoio": "#003C64",
    "fundo": "#F5F7FA",
    "card": "#FFFFFF",
    "borda": "#D9DDE3",
    "texto": "#1F2937",
    "texto_sec": "#6B7280",
    "sucesso": "#2E7D32",
    "erro": "#B91C1C",
    "alerta": "#B7791F",
    "branco": "#FFFFFF",
}


COLUNAS_BASE = [
    "EntryID",
    "Data Recebimento",
    "Hora Recebimento",
    "Dias Desde Recebimento",
    "Remetente Nome",
    "Remetente E-mail",
    "Para",
    "CC",
    "Assunto",
    "Corpo do E-mail",
    "Corpo Limpo",
    "Texto Truncado",
    "Tem Anexo",
    "Quantidade de Anexos",
    "Nomes dos Anexos",
    "Status Outlook",
    "Lido",
    "Importância Outlook",
    "Tamanho do Texto",
    "Data Extração",
    "Análise ChatGPT",
    "Urgência Sugerida",
    "Motivo da Urgência",
    "Categoria Sugerida",
    "Ação Recomendada",
    "Responsável Sugerido",
    "Observação RH",
    "Status Atendimento",
]


COLUNAS_GLOSSARIO = [
    "Palavra-chave",
    "Categoria",
    "Prioridade",
    "Ação sugerida",
    "Responsável",
    "Observação",
    "Ativo",
]


def garantir_pastas():
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    SAIDA_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)


def agora_texto():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def hoje_data():
    return datetime.now().date()


def limpar_texto(texto):
    if texto is None:
        return ""

    texto = str(texto)
    texto = texto.replace("\r", "\n")
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = texto.replace("\x00", "")
    return texto.strip()


def limpar_texto_uma_linha(texto):
    texto = limpar_texto(texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def limitar_excel(texto, limite=32000):
    texto = "" if texto is None else str(texto)
    if len(texto) <= limite:
        return texto, "Não"
    return texto[:limite] + "\n\n[TEXTO TRUNCADO PARA COMPATIBILIDADE COM EXCEL]", "Sim"


def normalizar_data_outlook(valor):
    if valor is None:
        return None

    try:
        if hasattr(valor, "date"):
            return valor
    except Exception:
        pass

    try:
        return datetime.fromtimestamp(time.mktime(valor.timetuple()))
    except Exception:
        return None


def dias_desde_recebimento(recebido):
    if not recebido:
        return ""

    try:
        data_email = recebido.date()
        return (hoje_data() - data_email).days
    except Exception:
        return ""


def traduzir_importancia(valor):
    try:
        valor = int(valor)
    except Exception:
        return ""

    mapa = {
        0: "Baixa",
        1: "Normal",
        2: "Alta",
    }
    return mapa.get(valor, str(valor))


def carregar_config():
    garantir_pastas()

    config_padrao = {
        "modo_caixa": "principal",
        "caixa_compartilhada": "",
        "pasta_outlook": "Caixa de Entrada",
        "quantidade": 200,
        "ultimos_dias": 30,
        "somente_nao_lidos": False,
        "pasta_saida": str(SAIDA_DIR),
        "nome_excel": DEFAULT_EXCEL_NAME,
    }

    if not CONFIG_FILE.exists():
        salvar_config(config_padrao)
        return config_padrao

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            config_lido = json.load(f)

        for chave, valor in config_padrao.items():
            config_lido.setdefault(chave, valor)

        return config_lido

    except Exception:
        return config_padrao


def salvar_config(config):
    garantir_pastas()
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def registrar_log_arquivo(mensagem):
    garantir_pastas()
    caminho_log = LOGS_DIR / "log_execucao.txt"
    with open(caminho_log, "a", encoding="utf-8") as f:
        f.write(f"[{agora_texto()}] {mensagem}\n")


def carregar_glossario(caminho_excel):
    """
    Função reservada para próxima iteração.

    Objetivo futuro:
    Ler a aba GLOSSARIO do Excel e usar as palavras-chave para sugerir
    categoria, prioridade, responsável e ação recomendada.

    Nesta primeira versão, a função não aplica classificação.
    """
    return []


def classificar_por_glossario(texto_email, glossario):
    """
    Função reservada para próxima iteração.

    Futuramente:
    - Receberá o texto do e-mail
    - Comparará com o glossário
    - Retornará categoria, prioridade, palavras encontradas,
      responsável e ação sugerida

    Nesta primeira versão:
    Retorna valores em branco ou Não classificado.
    """
    return {
        "analise_chatgpt": "",
        "urgencia": "",
        "motivo_urgencia": "",
        "categoria": "Não classificado",
        "acao_recomendada": "",
        "responsavel": "",
        "observacao": "",
        "status_atendimento": "Pendente",
    }


def get_smtp_address(mail):
    """
    Tenta obter o e-mail SMTP real do remetente.
    Funciona melhor em contas Exchange/Outlook corporativas.
    """
    try:
        if getattr(mail, "SenderEmailType", "") == "EX":
            try:
                exchange_user = mail.Sender.GetExchangeUser()
                if exchange_user:
                    smtp = exchange_user.PrimarySmtpAddress
                    if smtp:
                        return smtp
            except Exception:
                pass

            try:
                PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
                smtp = mail.Sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                if smtp:
                    return smtp
            except Exception:
                pass

        return getattr(mail, "SenderEmailAddress", "") or ""

    except Exception:
        return ""


def obter_pasta_por_nome(pasta_raiz, nome_pasta):
    """
    Procura uma pasta pelo nome dentro da árvore do Outlook.
    Se nome_pasta for vazio ou Caixa de Entrada, usa a própria pasta recebida.
    """
    nome_pasta = (nome_pasta or "").strip()

    if nome_pasta == "":
        return pasta_raiz

    nomes_padrao_inbox = {
        "caixa de entrada",
        "inbox",
        "entrada",
    }

    if nome_pasta.lower() in nomes_padrao_inbox:
        return pasta_raiz

    def buscar_recursivo(pasta):
        try:
            if pasta.Name.strip().lower() == nome_pasta.lower():
                return pasta
        except Exception:
            pass

        try:
            for subpasta in pasta.Folders:
                achou = buscar_recursivo(subpasta)
                if achou is not None:
                    return achou
        except Exception:
            pass

        return None

    encontrado = buscar_recursivo(pasta_raiz)

    if encontrado is None:
        raise Exception(f"Pasta não encontrada no Outlook: {nome_pasta}")

    return encontrado


def obter_pasta_outlook(modo_caixa, caixa_compartilhada, pasta_outlook, log_callback=None):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

    modo_caixa = (modo_caixa or "principal").strip().lower()
    caixa_compartilhada = (caixa_compartilhada or "").strip()
    pasta_outlook = (pasta_outlook or "Caixa de Entrada").strip()

    if modo_caixa == "compartilhada":
        if not caixa_compartilhada:
            raise Exception("Informe o nome ou e-mail da caixa compartilhada.")

        if log_callback:
            log_callback(f"Tentando acessar caixa compartilhada: {caixa_compartilhada}")

        recipient = outlook.CreateRecipient(caixa_compartilhada)
        recipient.Resolve()

        if not recipient.Resolved:
            raise Exception(f"Não foi possível resolver a caixa compartilhada: {caixa_compartilhada}")

        inbox = outlook.GetSharedDefaultFolder(recipient, 6)

    else:
        if log_callback:
            log_callback("Acessando Caixa de Entrada principal do Outlook.")
        inbox = outlook.GetDefaultFolder(6)

    pasta_final = obter_pasta_por_nome(inbox, pasta_outlook)

    if log_callback:
        try:
            log_callback(f"Pasta selecionada: {pasta_final.Name}")
        except Exception:
            log_callback("Pasta selecionada com sucesso.")

    return pasta_final


def extrair_dados_email(mail):
    try:
        entry_id = getattr(mail, "EntryID", "") or ""
    except Exception:
        entry_id = ""

    try:
        recebido = normalizar_data_outlook(getattr(mail, "ReceivedTime", None))
    except Exception:
        recebido = None

    data_recebimento = ""
    hora_recebimento = ""

    if recebido:
        try:
            data_recebimento = recebido.strftime("%d/%m/%Y")
            hora_recebimento = recebido.strftime("%H:%M:%S")
        except Exception:
            pass

    try:
        remetente_nome = getattr(mail, "SenderName", "") or ""
    except Exception:
        remetente_nome = ""

    remetente_email = get_smtp_address(mail)

    try:
        para = getattr(mail, "To", "") or ""
    except Exception:
        para = ""

    try:
        cc = getattr(mail, "CC", "") or ""
    except Exception:
        cc = ""

    try:
        assunto = getattr(mail, "Subject", "") or ""
    except Exception:
        assunto = ""

    try:
        corpo = getattr(mail, "Body", "") or ""
    except Exception:
        corpo = ""

    corpo = limpar_texto(corpo)
    corpo_limpo = limpar_texto_uma_linha(corpo)

    corpo_excel, truncado = limitar_excel(corpo)
    corpo_limpo_excel, truncado_limpo = limitar_excel(corpo_limpo)

    texto_truncado = "Sim" if truncado == "Sim" or truncado_limpo == "Sim" else "Não"

    nomes_anexos = []
    qtd_anexos = 0

    try:
        qtd_anexos = int(mail.Attachments.Count)
        for i in range(1, qtd_anexos + 1):
            try:
                nomes_anexos.append(mail.Attachments.Item(i).FileName)
            except Exception:
                nomes_anexos.append("Anexo não identificado")
    except Exception:
        qtd_anexos = 0

    tem_anexo = "Sim" if qtd_anexos > 0 else "Não"

    try:
        nao_lido = bool(getattr(mail, "UnRead", False))
        lido = "Não" if nao_lido else "Sim"
    except Exception:
        lido = ""

    try:
        importancia = traduzir_importancia(getattr(mail, "Importance", ""))
    except Exception:
        importancia = ""

    tamanho_texto = len(corpo)

    glossario = []
    classificacao = classificar_por_glossario(
        f"{assunto}\n{corpo}",
        glossario
    )

    linha = {
        "EntryID": entry_id,
        "Data Recebimento": data_recebimento,
        "Hora Recebimento": hora_recebimento,
        "Dias Desde Recebimento": dias_desde_recebimento(recebido),
        "Remetente Nome": remetente_nome,
        "Remetente E-mail": remetente_email,
        "Para": para,
        "CC": cc,
        "Assunto": assunto,
        "Corpo do E-mail": corpo_excel,
        "Corpo Limpo": corpo_limpo_excel,
        "Texto Truncado": texto_truncado,
        "Tem Anexo": tem_anexo,
        "Quantidade de Anexos": qtd_anexos,
        "Nomes dos Anexos": " | ".join(nomes_anexos),
        "Status Outlook": "Não lido" if lido == "Não" else "Lido",
        "Lido": lido,
        "Importância Outlook": importancia,
        "Tamanho do Texto": tamanho_texto,
        "Data Extração": agora_texto(),
        "Análise ChatGPT": classificacao["analise_chatgpt"],
        "Urgência Sugerida": classificacao["urgencia"],
        "Motivo da Urgência": classificacao["motivo_urgencia"],
        "Categoria Sugerida": classificacao["categoria"],
        "Ação Recomendada": classificacao["acao_recomendada"],
        "Responsável Sugerido": classificacao["responsavel"],
        "Observação RH": classificacao["observacao"],
        "Status Atendimento": classificacao["status_atendimento"],
    }

    return linha


def ler_emails_outlook(
    modo_caixa,
    caixa_compartilhada,
    pasta_outlook,
    quantidade,
    ultimos_dias,
    somente_nao_lidos,
    progress_callback=None,
    log_callback=None,
):
    pythoncom.CoInitialize()

    try:
        pasta = obter_pasta_outlook(
            modo_caixa=modo_caixa,
            caixa_compartilhada=caixa_compartilhada,
            pasta_outlook=pasta_outlook,
            log_callback=log_callback,
        )

        items = pasta.Items
        items.Sort("[ReceivedTime]", True)

        total_outlook = items.Count

        if log_callback:
            log_callback(f"Total aparente de itens na pasta: {total_outlook}")

        registros = []
        erros = []

        quantidade = int(quantidade)
        ultimos_dias = int(ultimos_dias)

        limite_data = None
        if ultimos_dias > 0:
            limite_data = hoje_data()
            if log_callback:
                log_callback(f"Filtro de período ativado: últimos {ultimos_dias} dias.")

        analisados = 0
        coletados = 0

        for item in items:
            if coletados >= quantidade:
                break

            analisados += 1

            try:
                item_class = getattr(item, "Class", None)
                if item_class != 43:
                    continue

                try:
                    recebido = normalizar_data_outlook(getattr(item, "ReceivedTime", None))
                except Exception:
                    recebido = None

                if ultimos_dias > 0 and recebido:
                    dias = dias_desde_recebimento(recebido)
                    if isinstance(dias, int) and dias > ultimos_dias:
                        continue

                if somente_nao_lidos:
                    try:
                        if not bool(getattr(item, "UnRead", False)):
                            continue
                    except Exception:
                        continue

                dados = extrair_dados_email(item)
                registros.append(dados)
                coletados += 1

                if progress_callback:
                    progress_callback(coletados, quantidade)

                if log_callback and coletados % 10 == 0:
                    log_callback(f"E-mails coletados: {coletados}")

            except Exception as e:
                erros.append(str(e))
                continue

        if log_callback:
            log_callback(f"Leitura finalizada. Itens analisados: {analisados}. E-mails coletados: {coletados}.")

        return registros, erros

    finally:
        pythoncom.CoUninitialize()


def aplicar_estilo_planilha(ws, tipo="base"):
    azul = CORES["azul"].replace("#", "")
    navy = CORES["navy"].replace("#", "")
    branco = "FFFFFF"
    cinza_claro = "F2F2F2"
    borda_cor = "D9DDE3"

    fill_header = PatternFill("solid", fgColor=navy)
    fill_sub = PatternFill("solid", fgColor=cinza_claro)
    font_header = Font(color=branco, bold=True)
    font_titulo = Font(color=navy, bold=True, size=14)
    border = Border(
        left=Side(style="thin", color=borda_cor),
        right=Side(style="thin", color=borda_cor),
        top=Side(style="thin", color=borda_cor),
        bottom=Side(style="thin", color=borda_cor),
    )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    larguras = {}

    if tipo == "base":
        larguras = {
            "A": 28,
            "B": 16,
            "C": 14,
            "D": 18,
            "E": 28,
            "F": 32,
            "G": 36,
            "H": 28,
            "I": 45,
            "J": 80,
            "K": 80,
            "L": 16,
            "M": 14,
            "N": 18,
            "O": 40,
            "P": 18,
            "Q": 12,
            "R": 18,
            "S": 18,
            "T": 22,
            "U": 35,
            "V": 18,
            "W": 35,
            "X": 22,
            "Y": 35,
            "Z": 24,
            "AA": 35,
            "AB": 18,
        }
    elif tipo == "glossario":
        larguras = {
            "A": 28,
            "B": 24,
            "C": 18,
            "D": 40,
            "E": 24,
            "F": 40,
            "G": 12,
        }
    elif tipo == "resumo":
        larguras = {
            "A": 35,
            "B": 35,
        }

    for col, largura in larguras.items():
        ws.column_dimensions[col].width = largura

    ws.row_dimensions[1].height = 28

    if tipo == "base":
        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 80


def criar_aba_base(wb, registros):
    ws = wb.active
    ws.title = "BASE_EMAILS"

    ws.append(COLUNAS_BASE)

    for registro in registros:
        ws.append([registro.get(coluna, "") for coluna in COLUNAS_BASE])

    aplicar_estilo_planilha(ws, tipo="base")


def criar_aba_glossario(wb):
    ws = wb.create_sheet("GLOSSARIO")
    ws.append(COLUNAS_GLOSSARIO)

    exemplos = [
        ["", "", "", "", "", "Preencher futuramente com termos validados pela operação.", "Sim"],
    ]

    for linha in exemplos:
        ws.append(linha)

    aplicar_estilo_planilha(ws, tipo="glossario")


def criar_aba_resumo(wb, registros):
    ws = wb.create_sheet("RESUMO")

    total = len(registros)
    total_anexo = sum(1 for r in registros if r.get("Tem Anexo") == "Sim")
    total_sem_anexo = total - total_anexo
    total_nao_lidos = sum(1 for r in registros if r.get("Lido") == "Não")
    total_lidos = sum(1 for r in registros if r.get("Lido") == "Sim")

    datas = []
    for r in registros:
        data_txt = r.get("Data Recebimento", "")
        try:
            datas.append(datetime.strptime(data_txt, "%d/%m/%Y").date())
        except Exception:
            pass

    data_inicial = min(datas).strftime("%d/%m/%Y") if datas else ""
    data_final = max(datas).strftime("%d/%m/%Y") if datas else ""

    linhas = [
        ["Indicador", "Valor"],
        ["Total de e-mails extraídos", total],
        ["Total com anexo", total_anexo],
        ["Total sem anexo", total_sem_anexo],
        ["Total não lidos", total_nao_lidos],
        ["Total lidos", total_lidos],
        ["Data inicial", data_inicial],
        ["Data final", data_final],
        ["Data da extração", agora_texto()],
        ["Aplicativo", f"{APP_NAME} v{APP_VERSION}"],
        ["Observação", "A classificação será feita posteriormente com apoio do ChatGPT e/ou glossário."],
    ]

    for linha in linhas:
        ws.append(linha)

    aplicar_estilo_planilha(ws, tipo="resumo")

    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=CORES["navy"].replace("#", ""))
        cell.font = Font(color="FFFFFF", bold=True)


def gerar_excel(registros, caminho_saida):
    caminho_saida = Path(caminho_saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    if caminho_saida.exists():
        try:
            os.rename(caminho_saida, caminho_saida)
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            novo_nome = caminho_saida.with_name(f"{caminho_saida.stem}_{timestamp}{caminho_saida.suffix}")
            caminho_saida = novo_nome

    wb = Workbook()
    criar_aba_base(wb, registros)
    criar_aba_glossario(wb)
    criar_aba_resumo(wb, registros)

    wb.save(caminho_saida)
    return caminho_saida


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        garantir_pastas()

        self.config_app = carregar_config()
        self.processo_rodando = False
        self.caminho_excel_gerado = None

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME)
        self.geometry("1180x760")
        self.minsize(1050, 680)
        self.configure(fg_color=CORES["fundo"])

        self.criar_variaveis()
        self.criar_layout()
        self.carregar_variaveis_config()

    def criar_variaveis(self):
        self.var_modo_caixa = ctk.StringVar(value="principal")
        self.var_caixa_compartilhada = ctk.StringVar(value="")
        self.var_pasta_outlook = ctk.StringVar(value="Caixa de Entrada")
        self.var_quantidade = ctk.StringVar(value="200")
        self.var_ultimos_dias = ctk.StringVar(value="30")
        self.var_somente_nao_lidos = ctk.BooleanVar(value=False)
        self.var_pasta_saida = ctk.StringVar(value=str(SAIDA_DIR))
        self.var_nome_excel = ctk.StringVar(value=DEFAULT_EXCEL_NAME)

    def carregar_variaveis_config(self):
        self.var_modo_caixa.set(self.config_app.get("modo_caixa", "principal"))
        self.var_caixa_compartilhada.set(self.config_app.get("caixa_compartilhada", ""))
        self.var_pasta_outlook.set(self.config_app.get("pasta_outlook", "Caixa de Entrada"))
        self.var_quantidade.set(str(self.config_app.get("quantidade", 200)))
        self.var_ultimos_dias.set(str(self.config_app.get("ultimos_dias", 30)))
        self.var_somente_nao_lidos.set(bool(self.config_app.get("somente_nao_lidos", False)))
        self.var_pasta_saida.set(self.config_app.get("pasta_saida", str(SAIDA_DIR)))
        self.var_nome_excel.set(self.config_app.get("nome_excel", DEFAULT_EXCEL_NAME))

    def obter_config_tela(self):
        return {
            "modo_caixa": self.var_modo_caixa.get(),
            "caixa_compartilhada": self.var_caixa_compartilhada.get().strip(),
            "pasta_outlook": self.var_pasta_outlook.get().strip(),
            "quantidade": int(self.var_quantidade.get()),
            "ultimos_dias": int(self.var_ultimos_dias.get()),
            "somente_nao_lidos": bool(self.var_somente_nao_lidos.get()),
            "pasta_saida": self.var_pasta_saida.get().strip(),
            "nome_excel": self.var_nome_excel.get().strip(),
        }

    def criar_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.criar_header()
        self.criar_cards_config()
        self.criar_area_log()
        self.criar_rodape()

    def criar_header(self):
        header = ctk.CTkFrame(self, fg_color=CORES["navy"], corner_radius=0, height=92)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        titulo = ctk.CTkLabel(
            header,
            text="Leitor Outlook RH",
            font=("Segoe UI", 26, "bold"),
            text_color=CORES["branco"],
            anchor="w",
        )
        titulo.grid(row=0, column=0, padx=28, pady=(18, 0), sticky="ew")

        subtitulo = ctk.CTkLabel(
            header,
            text="Exportação segura dos e-mails da caixa Sonova/Gente e Gestão para Excel, sem mover, excluir ou responder mensagens.",
            font=("Segoe UI", 14),
            text_color=CORES["branco"],
            anchor="w",
        )
        subtitulo.grid(row=1, column=0, padx=28, pady=(2, 18), sticky="ew")

    def criar_cards_config(self):
        area = ctk.CTkFrame(self, fg_color=CORES["fundo"], corner_radius=0)
        area.grid(row=1, column=0, sticky="ew", padx=22, pady=18)
        area.grid_columnconfigure(0, weight=1)
        area.grid_columnconfigure(1, weight=1)

        card_origem = ctk.CTkFrame(
            area,
            fg_color=CORES["card"],
            border_color=CORES["borda"],
            border_width=1,
            corner_radius=14,
        )
        card_origem.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        card_origem.grid_columnconfigure(1, weight=1)

        card_saida = ctk.CTkFrame(
            area,
            fg_color=CORES["card"],
            border_color=CORES["borda"],
            border_width=1,
            corner_radius=14,
        )
        card_saida.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        card_saida.grid_columnconfigure(1, weight=1)

        self.criar_card_origem(card_origem)
        self.criar_card_saida(card_saida)

    def criar_titulo_card(self, parent, texto):
        label = ctk.CTkLabel(
            parent,
            text=texto,
            font=("Segoe UI", 17, "bold"),
            text_color=CORES["navy"],
            anchor="w",
        )
        label.grid(row=0, column=0, columnspan=3, padx=18, pady=(16, 10), sticky="ew")

    def criar_label(self, parent, texto, row):
        label = ctk.CTkLabel(
            parent,
            text=texto,
            font=("Segoe UI", 13),
            text_color=CORES["texto"],
            anchor="w",
        )
        label.grid(row=row, column=0, padx=(18, 8), pady=7, sticky="w")
        return label

    def criar_entry(self, parent, variable, row, placeholder=""):
        entry = ctk.CTkEntry(
            parent,
            textvariable=variable,
            placeholder_text=placeholder,
            height=34,
            fg_color=CORES["branco"],
            border_color=CORES["borda"],
            text_color=CORES["texto"],
        )
        entry.grid(row=row, column=1, padx=(4, 18), pady=7, sticky="ew")
        return entry

    def criar_card_origem(self, parent):
        self.criar_titulo_card(parent, "Origem dos e-mails")

        radio_principal = ctk.CTkRadioButton(
            parent,
            text="Caixa principal",
            variable=self.var_modo_caixa,
            value="principal",
            fg_color=CORES["azul"],
            hover_color=CORES["azul_apoio"],
            text_color=CORES["texto"],
        )
        radio_principal.grid(row=1, column=0, padx=18, pady=5, sticky="w")

        radio_compartilhada = ctk.CTkRadioButton(
            parent,
            text="Caixa compartilhada",
            variable=self.var_modo_caixa,
            value="compartilhada",
            fg_color=CORES["azul"],
            hover_color=CORES["azul_apoio"],
            text_color=CORES["texto"],
        )
        radio_compartilhada.grid(row=1, column=1, padx=8, pady=5, sticky="w")

        self.criar_label(parent, "Caixa compartilhada", 2)
        self.criar_entry(
            parent,
            self.var_caixa_compartilhada,
            2,
            "Exemplo: genteegestao@empresa.com.br",
        )

        self.criar_label(parent, "Pasta Outlook", 3)
        self.criar_entry(
            parent,
            self.var_pasta_outlook,
            3,
            "Caixa de Entrada ou nome de subpasta",
        )

        self.criar_label(parent, "Quantidade máxima", 4)
        self.criar_entry(parent, self.var_quantidade, 4, "200")

        self.criar_label(parent, "Últimos dias", 5)
        self.criar_entry(parent, self.var_ultimos_dias, 5, "30. Use 0 para não filtrar")

        chk = ctk.CTkCheckBox(
            parent,
            text="Ler somente e-mails não lidos",
            variable=self.var_somente_nao_lidos,
            fg_color=CORES["azul"],
            hover_color=CORES["azul_apoio"],
            text_color=CORES["texto"],
        )
        chk.grid(row=6, column=0, columnspan=2, padx=18, pady=(8, 16), sticky="w")

    def criar_card_saida(self, parent):
        self.criar_titulo_card(parent, "Saída em Excel")

        self.criar_label(parent, "Pasta de saída", 1)

        frame_saida = ctk.CTkFrame(parent, fg_color="transparent")
        frame_saida.grid(row=1, column=1, padx=(4, 18), pady=7, sticky="ew")
        frame_saida.grid_columnconfigure(0, weight=1)

        entry_saida = ctk.CTkEntry(
            frame_saida,
            textvariable=self.var_pasta_saida,
            height=34,
            fg_color=CORES["branco"],
            border_color=CORES["borda"],
            text_color=CORES["texto"],
        )
        entry_saida.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        btn_pasta = ctk.CTkButton(
            frame_saida,
            text="Selecionar",
            command=self.selecionar_pasta_saida,
            width=95,
            height=34,
            fg_color=CORES["azul_apoio"],
            hover_color=CORES["navy"],
            text_color=CORES["branco"],
        )
        btn_pasta.grid(row=0, column=1)

        self.criar_label(parent, "Nome do Excel", 2)
        self.criar_entry(parent, self.var_nome_excel, 2, DEFAULT_EXCEL_NAME)

        info = ctk.CTkLabel(
            parent,
            text="O Excel será gerado com as abas BASE_EMAILS, GLOSSARIO e RESUMO. A classificação será feita depois, com análise da planilha.",
            font=("Segoe UI", 13),
            text_color=CORES["texto_sec"],
            justify="left",
            wraplength=480,
            anchor="w",
        )
        info.grid(row=3, column=0, columnspan=2, padx=18, pady=(8, 12), sticky="ew")

        botoes = ctk.CTkFrame(parent, fg_color="transparent")
        botoes.grid(row=4, column=0, columnspan=2, padx=18, pady=(4, 16), sticky="ew")
        botoes.grid_columnconfigure(0, weight=1)
        botoes.grid_columnconfigure(1, weight=1)

        self.btn_extrair = ctk.CTkButton(
            botoes,
            text="Extrair e-mails para Excel",
            command=self.iniciar_extracao,
            height=42,
            fg_color=CORES["azul"],
            hover_color=CORES["navy"],
            text_color=CORES["branco"],
            font=("Segoe UI", 14, "bold"),
        )
        self.btn_extrair.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.btn_abrir = ctk.CTkButton(
            botoes,
            text="Abrir pasta de saída",
            command=self.abrir_pasta_saida,
            height=42,
            fg_color=CORES["azul_apoio"],
            hover_color=CORES["navy"],
            text_color=CORES["branco"],
            font=("Segoe UI", 14, "bold"),
        )
        self.btn_abrir.grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def criar_area_log(self):
        area = ctk.CTkFrame(self, fg_color=CORES["fundo"], corner_radius=0)
        area.grid(row=2, column=0, sticky="nsew", padx=22, pady=(0, 12))
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(2, weight=1)

        card = ctk.CTkFrame(
            area,
            fg_color=CORES["card"],
            border_color=CORES["borda"],
            border_width=1,
            corner_radius=14,
        )
        card.grid(row=0, column=0, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        titulo = ctk.CTkLabel(
            card,
            text="Execução",
            font=("Segoe UI", 17, "bold"),
            text_color=CORES["navy"],
            anchor="w",
        )
        titulo.grid(row=0, column=0, padx=18, pady=(16, 8), sticky="ew")

        self.progress = ctk.CTkProgressBar(
            card,
            height=14,
            progress_color=CORES["azul"],
            fg_color=CORES["borda"],
        )
        self.progress.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="ew")
        self.progress.set(0)

        self.txt_log = ctk.CTkTextbox(
            card,
            height=280,
            fg_color="#0F172A",
            text_color="#E5E7EB",
            border_color=CORES["borda"],
            border_width=1,
            font=("Consolas", 12),
        )
        self.txt_log.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")

        self.log("Aplicativo iniciado.")
        self.log("Modo seguro: leitura e exportação. Nenhum e-mail será movido, excluído, respondido ou marcado.")

    def criar_rodape(self):
        rodape = ctk.CTkFrame(self, fg_color=CORES["fundo"], corner_radius=0, height=36)
        rodape.grid(row=3, column=0, sticky="ew", padx=22, pady=(0, 8))
        rodape.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(
            rodape,
            text="Anderson Marinho | Igarapé Digital",
            font=("Segoe UI", 12),
            text_color=CORES["texto_sec"],
            anchor="e",
        )
        label.grid(row=0, column=0, sticky="e")

    def log(self, mensagem):
        mensagem = str(mensagem)
        linha = f"[{agora_texto()}] {mensagem}\n"

        try:
            self.txt_log.insert("end", linha)
            self.txt_log.see("end")
        except Exception:
            pass

        try:
            registrar_log_arquivo(mensagem)
        except Exception:
            pass

    def selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(
            title="Selecione a pasta de saída",
            initialdir=self.var_pasta_saida.get() or str(SAIDA_DIR),
        )
        if pasta:
            self.var_pasta_saida.set(pasta)

    def abrir_pasta_saida(self):
        pasta = self.var_pasta_saida.get().strip()
        if not pasta:
            pasta = str(SAIDA_DIR)

        Path(pasta).mkdir(parents=True, exist_ok=True)
        os.startfile(pasta)

    def validar_campos(self):
        try:
            quantidade = int(self.var_quantidade.get())
            if quantidade <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Validação", "Informe uma quantidade válida maior que zero.")
            return False

        try:
            dias = int(self.var_ultimos_dias.get())
            if dias < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Validação", "Informe os últimos dias como número. Use 0 para não filtrar.")
            return False

        if self.var_modo_caixa.get() == "compartilhada" and not self.var_caixa_compartilhada.get().strip():
            messagebox.showerror("Validação", "Informe o nome ou e-mail da caixa compartilhada.")
            return False

        pasta_saida = self.var_pasta_saida.get().strip()
        if not pasta_saida:
            messagebox.showerror("Validação", "Informe a pasta de saída.")
            return False

        nome_excel = self.var_nome_excel.get().strip()
        if not nome_excel:
            messagebox.showerror("Validação", "Informe o nome do arquivo Excel.")
            return False

        if not nome_excel.lower().endswith(".xlsx"):
            self.var_nome_excel.set(nome_excel + ".xlsx")

        return True

    def travar_interface(self, travar=True):
        estado = "disabled" if travar else "normal"

        try:
            self.btn_extrair.configure(state=estado)
        except Exception:
            pass

        self.processo_rodando = travar

    def atualizar_progresso(self, atual, total):
        try:
            valor = min(max(atual / total, 0), 1)
            self.progress.set(valor)
        except Exception:
            pass

    def iniciar_extracao(self):
        if self.processo_rodando:
            messagebox.showinfo("Execução", "Já existe uma extração em andamento.")
            return

        if not self.validar_campos():
            return

        config = self.obter_config_tela()
        salvar_config(config)

        self.travar_interface(True)
        self.progress.set(0)
        self.log("Iniciando extração dos e-mails.")

        thread = threading.Thread(target=self.executar_extracao_thread, args=(config,), daemon=True)
        thread.start()

    def executar_extracao_thread(self, config):
        try:
            registros, erros = ler_emails_outlook(
                modo_caixa=config["modo_caixa"],
                caixa_compartilhada=config["caixa_compartilhada"],
                pasta_outlook=config["pasta_outlook"],
                quantidade=config["quantidade"],
                ultimos_dias=config["ultimos_dias"],
                somente_nao_lidos=config["somente_nao_lidos"],
                progress_callback=lambda atual, total: self.after(0, self.atualizar_progresso, atual, total),
                log_callback=lambda msg: self.after(0, self.log, msg),
            )

            if not registros:
                self.after(0, self.log, "Nenhum e-mail foi extraído com os filtros informados.")
                self.after(0, messagebox.showwarning, "Resultado", "Nenhum e-mail foi extraído com os filtros informados.")
                return

            pasta_saida = Path(config["pasta_saida"])
            nome_excel = config["nome_excel"]

            if not nome_excel.lower().endswith(".xlsx"):
                nome_excel += ".xlsx"

            caminho_saida = pasta_saida / nome_excel

            self.after(0, self.log, "Gerando arquivo Excel.")
            caminho_gerado = gerar_excel(registros, caminho_saida)

            self.caminho_excel_gerado = caminho_gerado

            self.after(0, self.progress.set, 1)
            self.after(0, self.log, f"Excel gerado com sucesso: {caminho_gerado}")

            if erros:
                self.after(0, self.log, f"Alguns itens não puderam ser lidos. Total de erros ignorados: {len(erros)}")

            self.after(
                0,
                messagebox.showinfo,
                "Concluído",
                f"Extração concluída com sucesso.\n\nE-mails extraídos: {len(registros)}\nArquivo:\n{caminho_gerado}",
            )

        except Exception as e:
            erro = traceback.format_exc()
            self.after(0, self.log, f"Erro na execução: {e}")
            registrar_log_arquivo(erro)
            self.after(0, messagebox.showerror, "Erro", f"Ocorreu um erro na execução:\n\n{e}")

        finally:
            self.after(0, self.travar_interface, False)


def main():
    garantir_pastas()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
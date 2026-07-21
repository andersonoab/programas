# -*- coding: utf-8 -*-
"""
Preenchedor SSA | v7.1 — Rateio + Dados Iniciais do Contas a Pagar
CustomerThink / Sonova OFC

Fluxo:
  1. Planilha  -> carrega o relatorio de folha (CSV/Excel) e mapeia colunas
  2. Verbas    -> escolhe a(s) rubrica(s) digitando ou clicando
  3. Rateio    -> agrega por CDC, calcula valor e percentual, fecha 100%
  4. SSA       -> preenche via F8 individual OU modo travado (loop automatico)
  5. Dados iniciais -> seleciona fornecedor/tipo no Excel e preenche via F8

Novidades v6.5:
  - Ao ativar Travar preenchimento, aplica automaticamente o preset seguro:
    0,08 entre acoes, 0,08 de delay inicial e 0,40 entre linhas.
  - Campo CDC tratado como controle de mascara fixa: nao usa Ctrl+A, Ctrl+C,
    Backspace ou Delete. Mantem o ponto da mascara e usa HOME para colocar o
    cursor no inicio antes de colar cada CDC.
  - Modo travado: loop automatico dos CDCs com ancora XY (clique de retorno
    ao primeiro campo do SSA antes de cada linha) para eliminar
    desalinhamento por popup, foco perdido ou TAB extra.
  - Pausa (F7 ou botao): interrompe o loop entre linhas, sem perder progresso.
  - Reposicionar ancora: pausa, aguarda F8 para gravar nova posicao e retoma.
  - Auto-pausa apos 2 falhas seguidas: nunca sai destruindo — pede intervencao.
  - F8 individual continua igual quando o modo travado esta desligado.

Novidades v7.1:
  - Nova aba "Dados iniciais" para carregar a Relacao SSA Fornecedores.
  - Selecao de fornecedor/tipo diretamente do Excel.
  - F8 campo a campo: Documento, Fornecedor e Tipo, sem TAB automatico entre eles.
  - Botao Voltar campo exclusivo da nova aba para recuperar a etapa anterior.
  - Documento recebe VALOR; Fornecedor recebe CODIGO; Tipo recebe
    CODIGO_DO_TIPO_DE_DOCUMENTO. TIPO_DE_DOCUMENTO e exibido pelas 3 primeiras
    letras para facilitar a conferencia.

Autor: Anderson Souza
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyautogui
import pyperclip
import time
import re
import json
import io
import threading
from pathlib import Path
from collections import defaultdict

try:
    from pynput import keyboard as pk
    HAS_PYNPUT = True
    PYNPUT_IMPORT_ERROR = ""
except Exception as e:
    pk = None
    HAS_PYNPUT = False
    PYNPUT_IMPORT_ERROR = str(e)

try:
    import keyboard as kb_global
    HAS_KEYBOARD_PACKAGE = True
    KEYBOARD_IMPORT_ERROR = ""
except Exception as e:
    kb_global = None
    HAS_KEYBOARD_PACKAGE = False
    KEYBOARD_IMPORT_ERROR = str(e)

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

pyautogui.PAUSE = 0.0
pyautogui.FAILSAFE = True


# =============================================================================
# NUCLEO DE DADOS (funcoes puras — testaveis sem GUI)
# =============================================================================

def parse_valor(v):
    """Converte texto de valor (BR ou bruto) em float. Trata 1.234,56 / 1,234.56 / 843.6."""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "nat"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in ("-", ".", ","):
        return None
    tem_v = "," in s
    tem_p = "." in s
    if tem_v and tem_p:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif tem_v:
        s = s.replace(",", ".")
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def fmt_brl(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def norm_cpf(v):
    return re.sub(r"\D", "", str(v))


def norm_cdc(v):
    """'424000 - 424000 AUDITIV MATRIZ' -> '424000'. '424000' -> '424000'."""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return ""
    m = re.match(r"^\s*([0-9A-Za-z\.]+)\s*[-–]\s*\S", s)
    if m:
        return m.group(1).strip()
    return s


def sem_acento(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", str(s))
                   if not unicodedata.combining(c))


def detectar_colunas(cols):
    """Autodeteccao das colunas do relatorio de folha. Retorna dict com as chaves canonicas."""
    low = {c: sem_acento(str(c).strip().lower()) for c in cols}
    achado = {"CPF": "", "NOME": "", "CDC": "", "VERBA": "", "DESCRICAO": "",
              "VALOR": "", "CLASSE": "", "EMPRESA": "", "MATRICULA": ""}

    for c, cl in low.items():
        if not achado["CPF"] and "cpf" in cl:
            achado["CPF"] = c
        if not achado["MATRICULA"] and ("matric" in cl or cl == "chapa"):
            achado["MATRICULA"] = c
        if not achado["NOME"] and cl == "nome":
            achado["NOME"] = c
        if not achado["CDC"] and ("centro de custo" in cl or cl in ("cdc", "c.c.", "cc")):
            achado["CDC"] = c
        if not achado["CLASSE"] and cl.startswith("clas"):
            achado["CLASSE"] = c
        if not achado["DESCRICAO"] and ("descri" in cl or "rubrica" in cl or "verba" in cl):
            achado["DESCRICAO"] = c
        if not achado["EMPRESA"] and ("estabelec" in cl or "empresa" in cl or "filial" in cl
                                      or cl in ("cnpj", "estab")):
            achado["EMPRESA"] = c

    if not achado["NOME"]:
        for c, cl in low.items():
            if "nome" in cl:
                achado["NOME"] = c
                break

    # Codigo da verba: coluna "Codigo" pura, sem qualificador de centro de custo
    for c, cl in low.items():
        base = cl.replace("ó", "o").replace("í", "i")
        if base in ("codigo", "cod", "cod.", "codigo verba", "codigo rubrica",
                    "codigo de folha", "codigo folha", "rubrica", "verba", "evento"):
            achado["VERBA"] = c
            break
    if not achado["VERBA"]:
        for c, cl in low.items():
            base = cl.replace("ó", "o")
            if "codigo" in base and "centro" not in base and "custo" not in base:
                achado["VERBA"] = c
                break

    # Valor: prioriza colunas do tipo "JUL/26 - Valor" e ignora data/hora
    candidatos = []
    for c, cl in low.items():
        if "dt " in cl or "data" in cl or "pgto" in cl or "hora" in cl:
            continue
        if "valor" in cl or cl.startswith("vl") or cl.startswith("vr") or cl == "total":
            candidatos.append(c)
    if candidatos:
        achado["VALOR"] = candidatos[-1]

    return achado


def montar_lancamentos(df, mapa, cadastro_cpf_cdc=None):
    """Transforma o DataFrame bruto em lista de lancamentos normalizados.

    mapa: dict com CPF/NOME/CDC/VERBA/DESCRICAO/VALOR/CLASSE (nomes de coluna, '' se ausente)
    cadastro_cpf_cdc: dict opcional {cpf_normalizado: cdc} usado quando a planilha nao traz CDC
    Retorna (lancamentos, sem_cdc) — sem_cdc = lista de lancamentos sem CDC resolvido.
    """
    lanc = []
    sem_cdc = []
    c_cpf = mapa.get("CPF") or ""
    c_nome = mapa.get("NOME") or ""
    c_cdc = mapa.get("CDC") or ""
    c_verba = mapa.get("VERBA") or ""
    c_desc = mapa.get("DESCRICAO") or ""
    c_valor = mapa.get("VALOR") or ""
    c_classe = mapa.get("CLASSE") or ""
    c_empresa = mapa.get("EMPRESA") or ""

    for i, row in df.iterrows():
        def g(col):
            if not col or col not in df.columns:
                return ""
            v = row[col]
            return "" if pd.isna(v) else str(v).strip()

        cpf = norm_cpf(g(c_cpf))
        cdc = norm_cdc(g(c_cdc))
        if not cdc and cadastro_cpf_cdc and cpf in cadastro_cpf_cdc:
            cdc = norm_cdc(cadastro_cpf_cdc[cpf])

        valor = parse_valor(g(c_valor))
        item = {
            "LINHA": int(i) + 2,
            "CPF": g(c_cpf),
            "CPF_N": cpf,
            "NOME": g(c_nome),
            "CDC": cdc,
            "VERBA": g(c_verba),
            "DESCRICAO": g(c_desc),
            "CLASSE": g(c_classe),
            "EMPRESA": g(c_empresa) or "(sem empresa)",
            "VALOR": valor if valor is not None else 0.0,
            "TEM_VALOR": valor is not None,
        }
        lanc.append(item)
        if not cdc:
            sem_cdc.append(item)
    return lanc, sem_cdc


def resumo_verbas(lancamentos):
    """Agrega por verba: codigo, descricao, classe, qtd, total, colaboradores."""
    agg = {}
    for l in lancamentos:
        k = l["VERBA"] or "(sem codigo)"
        if k not in agg:
            agg[k] = {"VERBA": k, "DESCRICAO": l["DESCRICAO"], "CLASSE": l["CLASSE"],
                      "QTD": 0, "TOTAL": 0.0, "CPFS": set(), "CDCS": set()}
        a = agg[k]
        a["QTD"] += 1
        a["TOTAL"] += l["VALOR"]
        if l["CPF_N"]:
            a["CPFS"].add(l["CPF_N"])
        if l["CDC"]:
            a["CDCS"].add(l["CDC"])
        if not a["DESCRICAO"] and l["DESCRICAO"]:
            a["DESCRICAO"] = l["DESCRICAO"]
    out = []
    for k, a in agg.items():
        a["PESSOAS"] = len(a["CPFS"])
        a["N_CDC"] = len(a["CDCS"])
        out.append(a)
    out.sort(key=lambda x: str(x["VERBA"]))
    return out


def calcular_rateio(lancamentos, verbas_sel, casas=2):
    """Agrega por CDC apenas as verbas selecionadas e calcula o percentual de rateio.

    Retorna (linhas, total). Cada linha: CDC, VALOR, PCT, QTD, PESSOAS, VERBAS.
    O percentual e arredondado e a sobra e ajustada no maior CDC para fechar 100,00%.
    """
    sel = set(str(v) for v in verbas_sel)
    agg = {}
    for l in lancamentos:
        if str(l["VERBA"]) not in sel:
            continue
        cdc = l["CDC"] or "(sem CDC)"
        if cdc not in agg:
            agg[cdc] = {"CDC": cdc, "VALOR": 0.0, "QTD": 0, "CPFS": set(), "VERBAS": set()}
        a = agg[cdc]
        a["VALOR"] += l["VALOR"]
        a["QTD"] += 1
        if l["CPF_N"]:
            a["CPFS"].add(l["CPF_N"])
        a["VERBAS"].add(str(l["VERBA"]))

    linhas = []
    for cdc, a in agg.items():
        linhas.append({"CDC": cdc, "VALOR": round(a["VALOR"], 2), "QTD": a["QTD"],
                       "PESSOAS": len(a["CPFS"]), "VERBAS": len(a["VERBAS"]), "PCT": 0.0})

    total = round(sum(l["VALOR"] for l in linhas), 2)
    if total:
        for l in linhas:
            l["PCT"] = round(l["VALOR"] / total * 100.0, casas)
        soma_pct = round(sum(l["PCT"] for l in linhas), casas)
        dif = round(100.0 - soma_pct, casas)
        if abs(dif) >= 10 ** (-casas) / 2:
            alvo = max(linhas, key=lambda x: abs(x["VALOR"]))
            alvo["PCT"] = round(alvo["PCT"] + dif, casas)

    linhas.sort(key=lambda x: -abs(x["VALOR"]))
    return linhas, total


def faixas_tercis(valores):
    v = [x for x in valores if x is not None]
    if len(v) < 3:
        return (0.0, 0.0)
    s = pd.Series(v)
    return (float(s.quantile(1 / 3)), float(s.quantile(2 / 3)))


def classe_faixa(valor, faixas):
    if valor is None:
        return "—"
    lo, hi = faixas
    if hi <= lo:
        return "—"
    if valor <= lo:
        return "Baixo"
    if valor <= hi:
        return "Medio"
    return "Alto"


def ler_csv_robusto(caminho):
    raw = open(caminho, "rb").read()
    for enc in ["utf-8-sig", "latin-1", "cp1252", "utf-8"]:
        for sep in [";", ",", "\t"]:
            try:
                df = pd.read_csv(io.BytesIO(raw), sep=sep, dtype=str, encoding=enc)
                if len(df.columns) > 1:
                    df.columns = [str(c).strip() for c in df.columns]
                    df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]
                    return df
            except Exception:
                continue
    raise ValueError("Nao foi possivel ler o CSV. Verifique o arquivo.")


def engine_para(caminho):
    cl = str(caminho).lower()
    if cl.endswith(".xlsx") or cl.endswith(".xlsm"):
        return "openpyxl"
    if cl.endswith(".xlsb"):
        return "pyxlsb"
    if cl.endswith(".xls"):
        return "xlrd"
    return None


def _limpar_df(df):
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed")]]
    return df.dropna(how="all")


def _engines_candidatos(caminho):
    """Ordem de tentativa: engine da extensao primeiro, depois os demais.
    Cobre arquivo com extensao trocada (xls que na verdade e xlsx, xlsx que e HTML, etc.)."""
    prim = engine_para(caminho)
    todos = ["openpyxl", "xlrd", "pyxlsb", "calamine", "odf"]
    if prim and prim in todos:
        todos.remove(prim)
        todos.insert(0, prim)
    return todos


def ler_excel(caminho, sheet_name=0):
    """Le xlsx/xlsm/xls/xlsb tentando cada engine. Se nao for Excel de verdade,
    cai para HTML (relatorios exportados como .xls) e depois para CSV."""
    erros = []
    for eng in _engines_candidatos(caminho):
        try:
            df = pd.read_excel(caminho, sheet_name=sheet_name, dtype=str, engine=eng)
            return _limpar_df(df)
        except ImportError as e:
            erros.append(f"{eng}: pacote ausente ({e})")
        except PermissionError:
            raise RuntimeError(f"O arquivo '{Path(caminho).name}' esta aberto no Excel. "
                               f"Feche o arquivo e tente novamente.")
        except Exception as e:
            erros.append(f"{eng}: {type(e).__name__}")
    # Fallback 1 — arquivo .xls que na verdade e HTML
    try:
        tabelas = pd.read_html(caminho, header=0)
        if tabelas:
            maior = max(tabelas, key=lambda t: t.shape[0] * t.shape[1])
            return _limpar_df(maior.astype(str))
    except Exception as e:
        erros.append(f"html: {type(e).__name__}")
    # Fallback 2 — texto delimitado com extensao de Excel
    try:
        return ler_csv_robusto(caminho)
    except Exception as e:
        erros.append(f"csv: {type(e).__name__}")
    raise RuntimeError(
        f"Nao foi possivel ler '{Path(caminho).name}'.\n\n"
        "Se for .xlsb instale: pip install pyxlsb\n"
        "Se for .xls antigo instale: pip install xlrd\n"
        "Se for .xlsx instale: pip install openpyxl\n\n"
        "Tentativas:\n  - " + "\n  - ".join(erros))


def abas_excel(caminho):
    """Lista as abas. Retorna [] quando o arquivo nao e um Excel real (HTML/CSV disfarcado)."""
    for eng in _engines_candidatos(caminho):
        try:
            return pd.ExcelFile(caminho, engine=eng).sheet_names
        except Exception:
            continue
    return []


def ler_tabela(caminho, aba=None):
    """Ponto unico de entrada: aceita csv, txt, xlsx, xlsm, xls, xlsb."""
    cl = str(caminho).lower()
    if cl.endswith(".csv") or cl.endswith(".txt"):
        return ler_csv_robusto(caminho)
    if aba is None or str(aba).startswith("("):
        return ler_excel(caminho)
    return ler_excel(caminho, sheet_name=aba)


def norm_cabecalho(v):
    """Normaliza cabecalhos mantendo uma forma comparavel e sem acentos."""
    s = sem_acento(str(v)).upper().strip()
    return re.sub(r"[^A-Z0-9]+", "_", s).strip("_")


def mapa_relacao_fornecedores(colunas):
    """Mapeia as colunas da Relacao SSA Fornecedores pelos nomes informados."""
    normalizadas = {norm_cabecalho(c): c for c in colunas}
    mapa = {
        "VALOR": normalizadas.get("VALOR", ""),
        "CODIGO": normalizadas.get("CODIGO", ""),
        "TIPO_DE_DOCUMENTO": normalizadas.get("TIPO_DE_DOCUMENTO", ""),
        "CODIGO_DO_TIPO_DE_DOCUMENTO": normalizadas.get("CODIGO_DO_TIPO_DE_DOCUMENTO", ""),
        "FORNECEDOR": "",
    }
    for candidato in (
            "FORNECEDOR", "NOME_DO_FORNECEDOR", "RAZAO_SOCIAL", "NOME_FANTASIA",
            "DESCRICAO_DO_FORNECEDOR", "DESCRICAO", "NOME"):
        if candidato in normalizadas:
            mapa["FORNECEDOR"] = normalizadas[candidato]
            break
    return mapa


def _ler_excel_sem_cabecalho(caminho, aba=None):
    erros = []
    for eng in _engines_candidatos(caminho):
        try:
            return pd.read_excel(caminho, sheet_name=aba if aba is not None else 0,
                                 header=None, dtype=str, engine=eng)
        except Exception as e:
            erros.append(f"{eng}: {type(e).__name__}")
    raise RuntimeError("Nao foi possivel procurar o cabecalho da relacao. " + ", ".join(erros))


def ler_relacao_fornecedores(caminho, aba=None):
    """Le a relacao mesmo quando o cabecalho nao esta na primeira linha.

    Retorna (df, mapa). CODIGO, TIPO_DE_DOCUMENTO e
    CODIGO_DO_TIPO_DE_DOCUMENTO sao obrigatorios. VALOR pode ser digitado na
    interface quando nao existir no arquivo.
    """
    df = ler_tabela(caminho, aba)
    mapa = mapa_relacao_fornecedores(df.columns)
    obrig = ("CODIGO", "TIPO_DE_DOCUMENTO", "CODIGO_DO_TIPO_DE_DOCUMENTO")
    if all(mapa[k] for k in obrig):
        return df.fillna(""), mapa

    cl = str(caminho).lower()
    if cl.endswith((".xlsx", ".xlsm", ".xls", ".xlsb")):
        bruto = _ler_excel_sem_cabecalho(caminho, aba)
        melhor_linha, melhor_pontos = None, -1
        for i in range(min(30, len(bruto))):
            nomes = [norm_cabecalho(v) for v in bruto.iloc[i].tolist()]
            pontos = sum(1 for k in obrig if k in nomes)
            if pontos > melhor_pontos:
                melhor_linha, melhor_pontos = i, pontos
        if melhor_linha is not None and melhor_pontos >= 2:
            cab = [str(v).strip() if not pd.isna(v) else f"COLUNA_{j + 1}"
                   for j, v in enumerate(bruto.iloc[melhor_linha].tolist())]
            df = bruto.iloc[melhor_linha + 1:].copy()
            df.columns = cab
            df = _limpar_df(df).fillna("")
            mapa = mapa_relacao_fornecedores(df.columns)

    faltantes = [k for k in obrig if not mapa.get(k)]
    if faltantes:
        encontrados = ", ".join(str(c) for c in df.columns)
        raise ValueError(
            "Colunas obrigatorias nao encontradas: " + ", ".join(faltantes) +
            "\n\nCabecalhos encontrados:\n" + encontrados)
    return df.fillna(""), mapa


# =============================================================================
# APLICACAO
# =============================================================================

class PreenchedorSSA:

    # ---- Paleta Sonova OFC (azul dominante 0083CA) ----
    AZUL = "#0083CA"          # dominante — acao primaria, header
    AZUL_HV = "#006BA8"
    AZUL_ESC = "#003C64"      # apoio escuro — titulos, KPI
    AZUL_ESC_HV = "#002A47"
    AZUL_CLARO = "#6EB4DC"    # apoio claro — acao secundaria
    AZUL_CLARO_HV = "#4E97C0"
    TEAL = "#005A64"
    TEAL_HV = "#00474F"
    PAINEL = "#F2F7FB"        # fundo de painel
    FUNDO = "#F4F6F9"
    CARD = "#FFFFFF"
    BORDA = "#CCCCCC"
    TEXTO = "#333333"
    TEXTO_SUAVE = "#646464"
    SUCESSO = "#1F7A3D"
    SUCESSO_HV = "#155E2E"
    VERDE_KPI = "#0F5C2E"
    ALERTA = "#8C321E"
    ALERTA_HV = "#6E2818"
    AMBAR = "#A97C00"
    AMBAR_HV = "#7A5500"
    NEUTRA = "#646464"
    NEUTRA_HV = "#4F4F4F"

    F_TIT = ("Segoe UI", 17, "bold")
    F_SUB = ("Segoe UI", 12)
    F_LBL = ("Segoe UI", 11, "bold")
    F_TXT = ("Segoe UI", 12)
    F_BTN = ("Segoe UI", 12, "bold")

    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Preenchedor SSA | v7.1 — Rateio + Dados Iniciais")
        self.app.geometry("1500x900")
        self.app.minsize(1120, 700)

        # ---------- Estado ----------
        self.arquivo = ""
        self.df_raw = None
        self.mapa = {}
        self.lancamentos = []
        self.verbas = []                 # resumo agregado por verba
        self.verbas_sel = set()          # codigos de verba selecionados
        self.rateio = []                 # linhas do rateio por CDC
        self.rateio_total = 0.0
        self.faixas = (0.0, 0.0)
        self.cadastro_map = {}           # cpf -> cdc (opcional)
        self.cadastro_nome = ""
        self._sort_verbas = ("VERBA", False)
        self._sort_rateio = ("VALOR", True)
        self.empresa_sel = "(Todas as empresas)"
        self._view_sort = {}          # ordenacao por clique nos cabecalhos
        self._head_txt = {}           # texto base de cada cabecalho

        # Estado SSA (compatibilidade v5)
        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.last_f8_time = 0.0
        self._debounce_secs = 0.30
        self._sending = False
        self.global_on = False
        self.hk_listener = None
        self._local_f8_bound = False

        # Estado do modo travado (loop automatico dos CDCs)
        self.anchor_xy = None            # (x, y) do primeiro campo do SSA
        self._locked = False             # modo travado ativo
        self._paused = False             # loop pausado
        self._awaiting_anchor = False    # aguardando F8 para capturar a ancora
        self._lock_thread = None         # thread do loop
        self._falhas_seguidas = 0        # auto-pausa apos 2 falhas seguidas

        # Estado dos dados iniciais do Contas a Pagar
        self.relacao_arquivo = ""
        self.relacao_df = None
        self.relacao_mapa = {}
        self.relacao_registros = []
        self.relacao_selecionado = None
        self.dados_iniciais_etapa = 1    # 1 = Documento; 2 = Fornecedor; 3 = Tipo; 4 = concluido

        self._estilos_ttk()
        self._construir()
        self._bind_local_keys()
        self.app.protocol("WM_DELETE_WINDOW", self.fechar_app)

        self.log("Aplicacao iniciada — v7.1 (rateio + dados iniciais do Contas a Pagar).")
        if HAS_PYNPUT:
            self.log("pynput disponivel: hotkey global habilitada.")
        elif HAS_KEYBOARD_PACKAGE:
            self.log("pacote keyboard disponivel: hotkey global habilitada.")
        else:
            self.log("pynput/keyboard ausentes: F8 funciona apenas com a janela do app em foco.")

    # =========================================================
    # BASE VISUAL
    # =========================================================
    def _estilos_ttk(self):
        st = ttk.Style()
        st.theme_use("default")
        st.configure("S.Treeview", background="white", foreground=self.TEXTO,
                     rowheight=30, fieldbackground="white", font=("Segoe UI", 11),
                     borderwidth=0)
        st.configure("S.Treeview.Heading", background="#E8F3FA", foreground=self.AZUL_ESC,
                     font=("Segoe UI", 11, "bold"), relief="flat", padding=(8, 8))
        st.map("S.Treeview",
               background=[("selected", "#BFE0F2")],
               foreground=[("selected", self.AZUL_ESC)])
        st.map("S.Treeview.Heading", background=[("active", "#D6EAF6")])

        # Scrollbars finas, sem setas, no tom Sonova
        for orient in ("Vertical", "Horizontal"):
            st.layout(f"S.{orient}.TScrollbar", [
                (f"{orient}.Scrollbar.trough", {
                    "sticky": "nswe",
                    "children": [(f"{orient}.Scrollbar.thumb",
                                  {"expand": "1", "sticky": "nswe"})]})])
            st.configure(f"S.{orient}.TScrollbar",
                         troughcolor="#EEF3F7", background="#B9CBD8",
                         bordercolor="#EEF3F7", darkcolor="#B9CBD8",
                         lightcolor="#B9CBD8", relief="flat", borderwidth=0,
                         arrowsize=0, width=11)
            st.map(f"S.{orient}.TScrollbar",
                   background=[("active", self.AZUL), ("pressed", self.AZUL_HV)])

    def card(self, parent):
        return ctk.CTkFrame(parent, fg_color=self.CARD, corner_radius=12,
                            border_width=1, border_color=self.BORDA)

    def titulo(self, parent, texto, sub=None):
        ctk.CTkLabel(parent, text=texto, font=self.F_TIT,
                     text_color=self.AZUL_ESC).pack(anchor="w", padx=16, pady=(14, 2))
        if sub:
            ctk.CTkLabel(parent, text=sub, font=self.F_SUB,
                         text_color=self.TEXTO_SUAVE).pack(anchor="w", padx=16, pady=(0, 10))

    def botao(self, parent, texto, comando, cor=None, hover=None, width=150, altura=38):
        return ctk.CTkButton(parent, text=texto, command=comando,
                             fg_color=cor or self.AZUL, hover_color=hover or self.AZUL_HV,
                             text_color="white", width=width, height=altura,
                             corner_radius=8, font=self.F_BTN)

    def entry(self, parent, width=None, placeholder=""):
        kw = {"height": 36, "fg_color": "white", "border_color": self.BORDA,
              "text_color": self.TEXTO, "font": self.F_TXT, "corner_radius": 8,
              "placeholder_text": placeholder}
        if width:
            kw["width"] = width
        return ctk.CTkEntry(parent, **kw)

    def combo(self, parent, values, command=None, width=None):
        kw = {"values": values, "command": command, "state": "readonly", "height": 36,
              "fg_color": "white", "border_color": self.BORDA, "button_color": self.AZUL,
              "button_hover_color": self.AZUL_HV, "text_color": self.TEXTO,
              "dropdown_fg_color": "white", "dropdown_text_color": self.TEXTO,
              "dropdown_hover_color": "#E8F3FA", "font": self.F_TXT, "corner_radius": 8}
        if width:
            kw["width"] = width
        return ctk.CTkComboBox(parent, **kw)

    def tabela(self, parent, cols_def, on_click=None, on_dclick=None, altura=None):
        """Treeview com moldura, barras finas que somem quando nao sao necessarias
        e colunas elasticas (a ultima coluna de texto absorve a sobra)."""
        wrap = ctk.CTkFrame(parent, fg_color="white", corner_radius=10,
                            border_width=1, border_color=self.BORDA)
        cols = [c[0] for c in cols_def]
        kw = {"columns": cols, "show": "headings", "style": "S.Treeview"}
        if altura:
            kw["height"] = altura
        tv = ttk.Treeview(wrap, **kw)
        base_txt = {}
        for col, txt, w, anc in cols_def:
            base_txt[col] = txt
            tv.heading(col, text=txt, command=lambda c=col, t=tv: self._sort_view(t, c))
            # colunas de texto (anchor w) esticam; numericas mantem largura
            tv.column(col, width=w, anchor=anc, minwidth=max(40, int(w * 0.6)),
                      stretch=(anc == "w"))
        self._head_txt[str(tv)] = base_txt
        tv.grid(row=0, column=0, sticky="nsew", padx=(2, 0), pady=(2, 0))

        sy = ttk.Scrollbar(wrap, orient="vertical", style="S.Vertical.TScrollbar",
                           command=tv.yview)
        sx = ttk.Scrollbar(wrap, orient="horizontal", style="S.Horizontal.TScrollbar",
                           command=tv.xview)

        def _auto_y(first, last):
            if float(first) <= 0.0 and float(last) >= 1.0:
                sy.grid_remove()
            else:
                sy.grid(row=0, column=1, sticky="ns", padx=(2, 2), pady=(2, 0))
            sy.set(first, last)

        def _auto_x(first, last):
            if float(first) <= 0.0 and float(last) >= 1.0:
                sx.grid_remove()
            else:
                sx.grid(row=1, column=0, sticky="ew", padx=(2, 0), pady=(2, 2))
            sx.set(first, last)

        tv.configure(yscrollcommand=_auto_y, xscrollcommand=_auto_x)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)
        tv.tag_configure("par", background="#FFFFFF")
        tv.tag_configure("impar", background="#F7FAFC")
        tv.tag_configure("sel", background="#E6F4EC", foreground="#155E2E")
        tv.tag_configure("sel_alt", background="#DCEFE3", foreground="#155E2E")
        tv.tag_configure("off", background="#F9F9F9", foreground="#8A8A8A")
        if on_click:
            tv.bind("<Button-1>", on_click)
        if on_dclick:
            tv.bind("<Double-1>", on_dclick)
        return wrap, tv

    def _rolavel(self, parent):
        """Envolve a aba num canvas: expande quando cabe, rola quando nao cabe.
        Garante que nenhuma tabela fique fora da area visivel."""
        parent.configure(fg_color=self.FUNDO)
        vbar = ttk.Scrollbar(parent, orient="vertical", style="S.Vertical.TScrollbar")
        cv = tk.Canvas(parent, bg=self.FUNDO, highlightthickness=0, bd=0,
                       yscrollcommand=vbar.set)
        vbar.configure(command=cv.yview)
        cv.pack(side="left", fill="both", expand=True)
        inner = ctk.CTkFrame(cv, fg_color=self.FUNDO, corner_radius=0)
        win = cv.create_window((0, 0), window=inner, anchor="nw")

        def _sync(_=None):
            cw, ch = cv.winfo_width(), cv.winfo_height()
            if cw < 2 or ch < 2:
                return
            rh = inner.winfo_reqheight()
            alt = max(rh, ch)
            cv.itemconfig(win, width=cw, height=alt)
            cv.configure(scrollregion=(0, 0, cw, alt))
            if rh > ch + 2:
                vbar.pack(side="right", fill="y")
            else:
                vbar.pack_forget()
                cv.yview_moveto(0)

        inner.bind("<Configure>", lambda e: cv.after_idle(_sync))
        cv.bind("<Configure>", lambda e: cv.after_idle(_sync))

        def _wheel(e):
            if cv.winfo_exists():
                cv.yview_scroll(int(-1 * (e.delta / 120)), "units")

        cv.bind("<Enter>", lambda e: cv.bind_all("<MouseWheel>", _wheel))
        cv.bind("<Leave>", lambda e: cv.unbind_all("<MouseWheel>"))
        return inner

    def _sort_view(self, tv, col=None):
        """Ordena as linhas visiveis de qualquer Treeview. Primeiro clique = ascendente,
        segundo = descendente. Numeros ordenam como numero; texto, como texto."""
        chave = str(tv)
        if col is not None:
            atual = self._view_sort.get(chave)
            rev = (not atual[1]) if (atual and atual[0] == col) else False
            self._view_sort[chave] = (col, rev)
        estado = self._view_sort.get(chave)
        if not estado:
            return
        col, rev = estado
        if col not in tv["columns"]:
            return

        itens = []
        for iid in tv.get_children(""):
            bruto = tv.set(iid, col)
            num = parse_valor(bruto)
            itens.append((0 if num is not None else 1,
                          num if num is not None else 0.0,
                          str(bruto).lower(), iid))
        itens.sort(key=lambda t: (t[0], t[1], t[2]), reverse=rev)
        for pos, (_, _, _, iid) in enumerate(itens):
            tv.move(iid, "", pos)

        self._zebra(tv)
        base = self._head_txt.get(chave, {})
        for c, txt in base.items():
            tv.heading(c, text=txt + ("   v" if rev else "   ^") if c == col else txt)

    def _zebra(self, tv):
        """Reaplica a alternancia de cor preservando o significado da tag (sel / off)."""
        for n, iid in enumerate(tv.get_children("")):
            tags = tv.item(iid, "tags") or ()
            if "sel" in tags or "sel_alt" in tags:
                nova = "sel" if n % 2 == 0 else "sel_alt"
            elif "off" in tags:
                nova = "off"
            elif any(t in tags for t in ("ok", "erro", "ign", "pend")):
                continue
            else:
                nova = "par" if n % 2 == 0 else "impar"
            tv.item(iid, tags=(nova,))

    def kpi(self, parent, titulo, altura=92):
        f = ctk.CTkFrame(parent, fg_color=self.AZUL_ESC, corner_radius=12, height=altura)
        f.pack_propagate(False)
        inner = ctk.CTkFrame(f, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")
        lt = ctk.CTkLabel(inner, text=titulo, font=("Segoe UI", 11, "bold"), text_color="#9EC8E8")
        lt.pack()
        lv = ctk.CTkLabel(inner, text="—", font=("Segoe UI", 30, "bold"), text_color="white")
        lv.pack()
        ld = ctk.CTkLabel(inner, text="aguardando dados", font=("Segoe UI", 11), text_color="#9EC8E8")
        ld.pack()
        return f, lv, ld

    # =========================================================
    # ESTRUTURA
    # =========================================================
    def _construir(self):
        self.app.configure(fg_color=self.FUNDO)

        # Header solido azul dominante
        header = ctk.CTkFrame(self.app, fg_color=self.AZUL, corner_radius=0, height=78)
        header.pack(fill="x")
        header.pack_propagate(False)

        esq = ctk.CTkFrame(header, fg_color="transparent")
        esq.pack(side="left", fill="both", expand=True, padx=22, pady=12)
        ctk.CTkLabel(esq, text="Preenchedor SSA — Rateio e Contas a Pagar",
                     font=("Segoe UI", 24, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(esq, text="Planilha  >  Verba  >  Rateio  >  SSA  |  Dados iniciais via F8",
                     font=("Segoe UI", 12), text_color="#D6EAF6").pack(anchor="w", pady=(2, 0))

        dir_ = ctk.CTkFrame(header, fg_color="transparent")
        dir_.pack(side="right", padx=22, pady=12)
        self.lbl_hotkey_mode = ctk.CTkLabel(dir_, text="Hotkey: Local",
                                            font=("Segoe UI", 12, "bold"), text_color="white")
        self.lbl_hotkey_mode.pack(anchor="e")
        self.lbl_passo = ctk.CTkLabel(dir_, text="Passo 1 de 5 — carregar planilha",
                                      font=("Segoe UI", 12), text_color="#D6EAF6")
        self.lbl_passo.pack(anchor="e", pady=(2, 0))

        # Barra de contexto — seletor de empresa (vale para todas as abas)
        ctxt = ctk.CTkFrame(self.app, fg_color=self.CARD, corner_radius=0, height=54,
                            border_width=0)
        ctxt.pack(fill="x")
        ctxt.pack_propagate(False)
        linha_ctx = ctk.CTkFrame(ctxt, fg_color="transparent")
        linha_ctx.pack(fill="both", expand=True, padx=14, pady=8)
        ctk.CTkLabel(linha_ctx, text="Empresa / estabelecimento:", font=self.F_LBL,
                     text_color=self.AZUL_ESC).pack(side="left", padx=(0, 8))
        self.cmb_empresa = self.combo(linha_ctx, ["(Todas as empresas)"],
                                      command=self._ao_mudar_empresa, width=520)
        self.cmb_empresa.set("(Todas as empresas)")
        self.cmb_empresa.pack(side="left", padx=(0, 12))
        self.lbl_empresa_info = ctk.CTkLabel(linha_ctx, text="carregue a planilha para listar as empresas",
                                             font=self.F_SUB, text_color=self.TEXTO_SUAVE)
        self.lbl_empresa_info.pack(side="left")

        self.tabs = ctk.CTkTabview(
            self.app, fg_color="transparent",
            segmented_button_fg_color="#E8F3FA",
            segmented_button_selected_color=self.AZUL,
            segmented_button_selected_hover_color=self.AZUL_HV,
            segmented_button_unselected_color="#E8F3FA",
            segmented_button_unselected_hover_color="#D6EAF6",
            text_color="white", text_color_disabled=self.TEXTO,
            command=self._ao_trocar_aba)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        self.T1 = "  1. Planilha  "
        self.T2 = "  2. Verbas  "
        self.T3 = "  3. Rateio  "
        self.T4 = "  4. Executar SSA  "
        self.T5 = "  5. Dados iniciais  "
        t1 = self.tabs.add(self.T1)
        t2 = self.tabs.add(self.T2)
        t3 = self.tabs.add(self.T3)
        t4 = self.tabs.add(self.T4)
        t5 = self.tabs.add(self.T5)

        self.aba_planilha(t1)
        self.aba_verbas(t2)
        self.aba_rateio(t3)
        self.aba_ssa(t4)
        self.aba_dados_iniciais(t5)

        self.atualizar_resumo()
        self.atualizar_linha_atual()

    def _ao_trocar_aba(self):
        atual = self.tabs.get()
        mapa = {self.T1: "Passo 1 de 5 — carregar planilha",
                self.T2: "Passo 2 de 5 — escolher a verba",
                self.T3: "Passo 3 de 5 — conferir o rateio",
                self.T4: "Passo 4 de 5 — preencher o rateio no SSA",
                self.T5: "Passo 5 de 5 — preencher os dados iniciais"}
        self.lbl_passo.configure(text=mapa.get(atual, ""))

    # =========================================================
    # ABA 1 — PLANILHA
    # =========================================================
    def aba_planilha(self, parent):
        parent = self._rolavel(parent)
        parent.grid_columnconfigure(0, weight=3)
        parent.grid_columnconfigure(1, weight=2)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(1, weight=1, minsize=260)

        # --- Card: arquivo + mapeamento ---
        c = self.card(parent)
        c.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 8))
        self.titulo(c, "Planilha de folha",
                    "Relatorio de Valores por Codigo de Folha (CSV ou Excel). As colunas sao detectadas automaticamente.")

        lin = ctk.CTkFrame(c, fg_color="transparent")
        lin.pack(fill="x", padx=16, pady=(0, 8))
        self.botao(lin, "Carregar planilha", self.carregar_planilha, self.AZUL, self.AZUL_HV, 165).pack(side="left", padx=(0, 8))
        self.cmb_aba = self.combo(lin, [], command=self._ao_mudar_aba_excel, width=190)
        self.cmb_aba.pack(side="left", padx=(0, 8))
        self.botao(lin, "Nova planilha", self.resetar_planilha, self.NEUTRA, self.NEUTRA_HV, 130).pack(side="left")

        self.lbl_arquivo = ctk.CTkLabel(c, text="Nenhum arquivo carregado.", font=self.F_SUB,
                                        text_color=self.TEXTO_SUAVE)
        self.lbl_arquivo.pack(anchor="w", padx=16, pady=(0, 10))

        box = ctk.CTkFrame(c, fg_color=self.PAINEL, corner_radius=10,
                           border_width=1, border_color=self.BORDA)
        box.pack(fill="x", padx=16, pady=(0, 12))
        ctk.CTkLabel(box, text="Mapeamento de colunas", font=("Segoe UI", 13, "bold"),
                     text_color=self.AZUL_ESC).pack(anchor="w", padx=12, pady=(10, 8))

        g = ctk.CTkFrame(box, fg_color="transparent")
        g.pack(fill="x", padx=12, pady=(0, 10))
        for i in range(5):
            g.grid_columnconfigure(i, weight=1)

        campos = [("Codigo da verba", "VERBA"), ("Descricao da verba", "DESCRICAO"),
                  ("Valor", "VALOR"), ("CDC (centro de custo)", "CDC"),
                  ("Empresa / estabelecimento", "EMPRESA"), ("CPF", "CPF"),
                  ("Nome", "NOME"), ("Classificacao", "CLASSE"),
                  ("Matricula", "MATRICULA"), ("", "")]
        self.cmb_map = {}
        for i, (rot, chave) in enumerate(campos):
            if not chave:
                continue
            r, col = divmod(i, 5)
            ctk.CTkLabel(g, text=rot, font=("Segoe UI", 10, "bold"), text_color=self.TEXTO
                         ).grid(row=r * 2, column=col, sticky="w", padx=(0, 6), pady=(4, 2))
            cb = self.combo(g, ["(nenhuma)"], command=lambda v, k=chave: self._map_alterado(k, v))
            cb.configure(height=30, font=("Segoe UI", 11))
            cb.set("(nenhuma)")
            cb.grid(row=r * 2 + 1, column=col, sticky="ew", padx=(0, 6), pady=(0, 2))
            self.cmb_map[chave] = cb

        acoes = ctk.CTkFrame(c, fg_color="transparent")
        acoes.pack(fill="x", padx=16, pady=(0, 14))
        self.botao(acoes, "Processar planilha", self.processar_planilha,
                   self.SUCESSO, self.SUCESSO_HV, 175).pack(side="left", padx=(0, 8))
        self.botao(acoes, "Salvar perfil", self.salvar_perfil, self.AZUL_CLARO, self.AZUL_CLARO_HV, 130).pack(side="left", padx=(0, 8))
        self.botao(acoes, "Carregar perfil", self.carregar_perfil, self.AZUL_CLARO, self.AZUL_CLARO_HV, 145).pack(side="left")

        # --- Card: cadastro opcional ---
        c2 = self.card(parent)
        c2.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 8))
        self.titulo(c2, "Cadastro (opcional)",
                    "Use apenas se a planilha NAO trouxer o CDC. O CPF sera cruzado com o cadastro para trazer o centro de custo.")

        lin2 = ctk.CTkFrame(c2, fg_color="transparent")
        lin2.pack(fill="x", padx=16, pady=(0, 8))
        self.botao(lin2, "Carregar cadastro", self.carregar_cadastro,
                   self.AZUL_ESC, self.AZUL_ESC_HV, 165).pack(side="left", padx=(0, 8))
        self.botao(lin2, "Remover", self.remover_cadastro, self.NEUTRA, self.NEUTRA_HV, 100).pack(side="left")

        self.lbl_cadastro = ctk.CTkLabel(c2, text="Cadastro nao carregado.", font=self.F_SUB,
                                         text_color=self.TEXTO_SUAVE)
        self.lbl_cadastro.pack(anchor="w", padx=16, pady=(0, 8))

        gg = ctk.CTkFrame(c2, fg_color=self.PAINEL, corner_radius=10,
                          border_width=1, border_color=self.BORDA)
        gg.pack(fill="x", padx=16, pady=(0, 12))
        gg.grid_columnconfigure(0, weight=1)
        gg.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(gg, text="Coluna CPF", font=self.F_LBL, text_color=self.TEXTO
                     ).grid(row=0, column=0, sticky="w", padx=(12, 6), pady=(10, 3))
        ctk.CTkLabel(gg, text="Coluna CDC", font=self.F_LBL, text_color=self.TEXTO
                     ).grid(row=0, column=1, sticky="w", padx=(6, 12), pady=(10, 3))
        self.cmb_cad_cpf = self.combo(gg, [])
        self.cmb_cad_cpf.grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=(0, 12))
        self.cmb_cad_cdc = self.combo(gg, [])
        self.cmb_cad_cdc.grid(row=1, column=1, sticky="ew", padx=(6, 12), pady=(0, 12))

        self.lbl_diag = ctk.CTkLabel(c2, text="", font=("Segoe UI", 11), justify="left",
                                     text_color=self.TEXTO_SUAVE)
        self.lbl_diag.pack(anchor="w", padx=16, pady=(0, 14))

        # --- Card: preview ---
        c3 = self.card(parent)
        c3.grid(row=1, column=0, columnspan=2, sticky="nsew")
        self.titulo(c3, "Previa dos lancamentos",
                    "Clique no cabecalho para ordenar. Mostra ate 600 linhas da empresa selecionada.")
        cols = [("LINHA", "Linha", 60, "center"), ("NOME", "Nome", 240, "w"),
                ("CPF", "CPF", 120, "center"), ("CDC", "CDC", 90, "center"),
                ("VERBA", "Codigo", 80, "center"), ("DESCRICAO", "Descricao da verba", 260, "w"),
                ("CLASSE", "Clas.", 70, "center"), ("VALOR", "Valor", 120, "e")]
        wrap, self.tv_preview = self.tabela(c3, cols, altura=9)
        wrap.pack(fill="both", expand=True, padx=16, pady=(0, 14))

    def _map_alterado(self, chave, valor):
        self.mapa[chave] = "" if valor in ("", "(nenhuma)") else valor

    def carregar_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar planilha de folha",
            filetypes=[("Todos suportados", "*.xlsx *.xlsm *.xls *.xlsb *.csv *.txt *.ods"),
                       ("Excel", "*.xlsx *.xlsm *.xls *.xlsb"), ("Texto", "*.csv *.txt"),
                       ("Todos", "*.*")])
        if not caminho:
            return
        try:
            self.arquivo = caminho
            cl = caminho.lower()
            if cl.endswith(".csv") or cl.endswith(".txt"):
                self.cmb_aba.configure(values=["(arquivo sem abas)"])
                self.cmb_aba.set("(arquivo sem abas)")
                df = ler_tabela(caminho)
            else:
                abas = abas_excel(caminho)
                if abas:
                    self.cmb_aba.configure(values=abas)
                    self.cmb_aba.set(abas[0])
                    df = ler_tabela(caminho, abas[0])
                else:
                    # .xls exportado como HTML/texto: nao tem abas reais
                    self.cmb_aba.configure(values=["(arquivo sem abas)"])
                    self.cmb_aba.set("(arquivo sem abas)")
                    df = ler_tabela(caminho)
            if df is None or df.empty:
                raise RuntimeError("O arquivo foi lido, mas nao contem linhas de dados.")
            self._aplicar_df(df)
            self.lbl_arquivo.configure(text=f"{Path(caminho).name}  |  {len(df)} linhas  |  {len(df.columns)} colunas",
                                       text_color=self.SUCESSO)
            self.log(f"Planilha carregada: {Path(caminho).name} ({len(df)} linhas).")
        except Exception as e:
            messagebox.showerror("Erro ao carregar planilha", str(e))

    def _ao_mudar_aba_excel(self, aba):
        if not self.arquivo or aba.startswith("("):
            return
        try:
            df = ler_tabela(self.arquivo, aba)
            self._aplicar_df(df)
            self.log(f"Aba trocada: {aba} ({len(df)} linhas).")
        except Exception as e:
            messagebox.showerror("Erro ao ler aba", str(e))

    def _aplicar_df(self, df):
        self.df_raw = df
        cols = list(df.columns)
        det = detectar_colunas(cols)
        self.mapa = dict(det)
        for chave, cb in self.cmb_map.items():
            cb.configure(values=["(nenhuma)"] + cols)
            cb.set(det.get(chave) or "(nenhuma)")
        achou = [k for k, v in det.items() if v]
        self.lbl_diag.configure(
            text=f"Autodeteccao: {len(achou)} de {len(det)} campos identificados.\n"
                 f"Ajuste manualmente o que estiver errado e clique em Processar planilha.")
        self.processar_planilha(silencioso=True)

    def carregar_cadastro(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar cadastro (SRA_Ativos)",
            filetypes=[("Todos suportados", "*.xlsx *.xlsm *.xls *.xlsb *.csv *.txt *.ods"),
                       ("Excel", "*.xlsx *.xlsm *.xls *.xlsb"), ("Texto", "*.csv *.txt"),
                       ("Todos", "*.*")])
        if not caminho:
            return
        try:
            df = ler_tabela(caminho)
            self._cad_df = df
            cols = list(df.columns)
            self.cmb_cad_cpf.configure(values=cols)
            self.cmb_cad_cdc.configure(values=cols)
            for c in cols:
                cl = str(c).lower()
                if "cpf" in cl:
                    self.cmb_cad_cpf.set(c)
                if "centro de custo" in cl or cl in ("cdc", "cc"):
                    self.cmb_cad_cdc.set(c)
            self.cadastro_nome = Path(caminho).name
            self.lbl_cadastro.configure(text=f"{self.cadastro_nome}  |  {len(df)} registros",
                                        text_color=self.SUCESSO)
            self._montar_cadastro_map()
            self.log(f"Cadastro carregado: {self.cadastro_nome} ({len(df)} registros).")
        except Exception as e:
            messagebox.showerror("Erro ao carregar cadastro", str(e))

    def _montar_cadastro_map(self):
        self.cadastro_map = {}
        df = getattr(self, "_cad_df", None)
        if df is None:
            return
        ccpf, ccdc = self.cmb_cad_cpf.get(), self.cmb_cad_cdc.get()
        if not ccpf or not ccdc or ccpf not in df.columns or ccdc not in df.columns:
            return
        for _, r in df.iterrows():
            k = norm_cpf(r[ccpf])
            if k:
                self.cadastro_map[k] = norm_cdc(r[ccdc])

    def remover_cadastro(self):
        self._cad_df = None
        self.cadastro_map = {}
        self.cadastro_nome = ""
        self.cmb_cad_cpf.configure(values=[])
        self.cmb_cad_cpf.set("")
        self.cmb_cad_cdc.configure(values=[])
        self.cmb_cad_cdc.set("")
        self.lbl_cadastro.configure(text="Cadastro nao carregado.", text_color=self.TEXTO_SUAVE)
        self.log("Cadastro removido.")

    def processar_planilha(self, silencioso=False):
        if self.df_raw is None:
            if not silencioso:
                messagebox.showwarning("Planilha", "Carregue a planilha primeiro.")
            return
        if not self.mapa.get("VERBA") or not self.mapa.get("VALOR"):
            if not silencioso:
                messagebox.showwarning("Mapeamento",
                                       "Indique ao menos a coluna do CODIGO DA VERBA e a coluna de VALOR.")
            return
        self._montar_cadastro_map()
        self.lancamentos, sem_cdc = montar_lancamentos(self.df_raw, self.mapa, self.cadastro_map)
        self._popular_empresas()
        self.verbas = resumo_verbas(self._lanc_ativos())
        self.verbas_sel = set()
        self.rateio, self.rateio_total = [], 0.0

        self._render_preview()
        self._render_verbas()
        self._render_rateio()

        total_geral = sum(l["VALOR"] for l in self.lancamentos)
        aviso = ""
        if sem_cdc:
            aviso = f"\nAtencao: {len(sem_cdc)} lancamento(s) sem CDC. Carregue o cadastro ou revise o mapeamento."
        self.lbl_diag.configure(
            text=f"{len(self.lancamentos)} lancamentos  |  {len(self.verbas)} verbas distintas  |  "
                 f"total do arquivo: {fmt_brl(total_geral)}{aviso}",
            text_color=self.ALERTA if sem_cdc else self.TEXTO_SUAVE)
        self.log(f"Processado: {len(self.lancamentos)} lancamentos, {len(self.verbas)} verbas.")
        if not silencioso:
            self.tabs.set(self.T2)
            self._ao_trocar_aba()

    # ---------- Filtro global por empresa ----------
    def _lanc_ativos(self):
        """Lancamentos da empresa selecionada (ou todos)."""
        if self.empresa_sel.startswith("(Todas"):
            return self.lancamentos
        return [l for l in self.lancamentos if l.get("EMPRESA") == self.empresa_sel]

    def _popular_empresas(self):
        agg = {}
        for l in self.lancamentos:
            e = l.get("EMPRESA") or "(sem empresa)"
            agg[e] = agg.get(e, 0) + 1
        opcoes = ["(Todas as empresas)"] + sorted(agg.keys())
        self.cmb_empresa.configure(values=opcoes)
        if self.empresa_sel not in opcoes:
            self.empresa_sel = "(Todas as empresas)"
        self.cmb_empresa.set(self.empresa_sel)
        self.lbl_empresa_info.configure(
            text=f"{len(agg)} empresa(s) na planilha  |  {len(self._lanc_ativos())} lancamento(s) em uso")

    def _ao_mudar_empresa(self, valor):
        self.empresa_sel = valor
        ativos = self._lanc_ativos()
        self.verbas = resumo_verbas(ativos)
        self.verbas_sel = set()
        self.rateio, self.rateio_total = [], 0.0
        self.lbl_empresa_info.configure(
            text=f"{len(ativos)} lancamento(s) em uso  |  total {fmt_brl(sum(l['VALOR'] for l in ativos))}")
        self._render_preview()
        self._render_verbas()
        self._render_rateio()
        self.log(f"Empresa selecionada: {valor} ({len(ativos)} lancamentos).")

    def _render_preview(self):
        for i in self.tv_preview.get_children():
            self.tv_preview.delete(i)
        for n, l in enumerate(self._lanc_ativos()[:600]):
            tag = "par" if n % 2 == 0 else "impar"
            self.tv_preview.insert("", "end", tags=(tag,),
                                   values=(l["LINHA"], l["NOME"], l["CPF"], l["CDC"], l["VERBA"],
                                           l["DESCRICAO"], l["CLASSE"], fmt_num(l["VALOR"])))
        self._sort_view(self.tv_preview)

    def resetar_planilha(self):
        if not messagebox.askyesno("Nova planilha",
                                   "Isso limpa a planilha, as verbas selecionadas e o rateio.\nDeseja continuar?"):
            return
        self.arquivo = ""
        self.df_raw = None
        self.lancamentos = []
        self.verbas = []
        self.verbas_sel = set()
        self.rateio, self.rateio_total = [], 0.0
        self.cmb_aba.configure(values=[])
        self.cmb_aba.set("")
        for cb in self.cmb_map.values():
            cb.configure(values=["(nenhuma)"])
            cb.set("(nenhuma)")
        self.empresa_sel = "(Todas as empresas)"
        self.cmb_empresa.configure(values=["(Todas as empresas)"])
        self.cmb_empresa.set("(Todas as empresas)")
        self.lbl_empresa_info.configure(text="carregue a planilha para listar as empresas")
        self.lbl_arquivo.configure(text="Nenhum arquivo carregado.", text_color=self.TEXTO_SUAVE)
        self.lbl_diag.configure(text="", text_color=self.TEXTO_SUAVE)
        self._render_preview()
        self._render_verbas()
        self._render_rateio()
        self.tabs.set(self.T1)
        self._ao_trocar_aba()
        self.log("Planilha e rateio zerados.")

    # =========================================================
    # ABA 2 — VERBAS
    # =========================================================
    def aba_verbas(self, parent):
        parent = self._rolavel(parent)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(1, weight=1, minsize=300)

        top = self.card(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.titulo(top, "Selecione a verba do rateio",
                    "Digite o codigo (ex.: 9951) ou parte da descricao (ex.: liquido). Clique na linha para marcar ou desmarcar.")

        bar = ctk.CTkFrame(top, fg_color=self.PAINEL, corner_radius=10,
                           border_width=1, border_color=self.BORDA)
        bar.pack(fill="x", padx=16, pady=(0, 10))

        l1 = ctk.CTkFrame(bar, fg_color="transparent")
        l1.pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(l1, text="Buscar verba:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(0, 8))
        self.ent_busca_verba = self.entry(l1, width=340, placeholder="codigo ou descricao...")
        self.ent_busca_verba.pack(side="left", padx=(0, 8))
        self.ent_busca_verba.bind("<KeyRelease>", lambda e: self._render_verbas())
        self.ent_busca_verba.bind("<Return>", self._enter_busca_verba)

        ctk.CTkLabel(l1, text="Clas.:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(8, 6))
        self.cmb_classe = self.combo(l1, ["Todas"], command=lambda v: self._render_verbas(), width=130)
        self.cmb_classe.set("Todas")
        self.cmb_classe.pack(side="left", padx=(0, 8))

        ctk.CTkLabel(l1, text="Ordenar por:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(8, 6))
        self.cmb_ord_verba = self.combo(l1, ["Codigo", "Maior valor", "Descricao", "Qtd lancamentos"],
                                        command=self._ordenar_verbas, width=170)
        self.cmb_ord_verba.set("Codigo")
        self.cmb_ord_verba.pack(side="left")

        l2 = ctk.CTkFrame(bar, fg_color="transparent")
        l2.pack(fill="x", padx=12, pady=(0, 10))
        self.botao(l2, "Marcar filtradas", lambda: self._marcar_verbas(True), self.AZUL, self.AZUL_HV, 150).pack(side="left", padx=(0, 8))
        self.botao(l2, "Desmarcar filtradas", lambda: self._marcar_verbas(False), self.AZUL_CLARO, self.AZUL_CLARO_HV, 165).pack(side="left", padx=(0, 8))
        self.botao(l2, "Limpar selecao", self._limpar_verbas, self.NEUTRA, self.NEUTRA_HV, 140).pack(side="left", padx=(0, 8))
        self.botao(l2, "Limpar busca", self._limpar_busca_verba, self.NEUTRA, self.NEUTRA_HV, 130).pack(side="left", padx=(0, 16))
        self.botao(l2, "Gerar rateio  >", self.gerar_rateio, self.SUCESSO, self.SUCESSO_HV, 175).pack(side="right")

        kf = ctk.CTkFrame(top, fg_color="transparent")
        kf.pack(fill="x", padx=16, pady=(0, 14))
        kf.grid_columnconfigure(0, weight=2)
        kf.grid_columnconfigure(1, weight=1)
        k1, self.kv_verba_total, self.kd_verba_total = self.kpi(kf, "TOTAL DAS VERBAS SELECIONADAS")
        k1.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        k2, self.kv_verba_qtd, self.kd_verba_qtd = self.kpi(kf, "VERBAS SELECIONADAS")
        k2.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        body = self.card(parent)
        body.grid(row=1, column=0, sticky="nsew")
        cols = [("SEL", "Sel.", 55, "center"), ("VERBA", "Codigo", 90, "center"),
                ("DESCRICAO", "Descricao da verba", 340, "w"), ("CLASSE", "Clas.", 80, "center"),
                ("QTD", "Lancamentos", 100, "center"), ("PESSOAS", "Pessoas", 85, "center"),
                ("N_CDC", "CDCs", 70, "center"), ("TOTAL", "Total", 140, "e")]
        wrap, self.tv_verbas = self.tabela(body, cols, on_click=self._clique_verba,
                                           on_dclick=self._dclique_verba, altura=10)
        wrap.pack(fill="both", expand=True, padx=16, pady=(14, 6))
        ctk.CTkLabel(body,
                     text="Enter na busca marca a primeira verba da lista. Duplo clique abre os lancamentos da verba.",
                     font=("Segoe UI", 10), text_color=self.TEXTO_SUAVE
                     ).pack(anchor="w", padx=16, pady=(0, 12))

    def _verbas_filtradas(self):
        q = self.ent_busca_verba.get().strip().lower() if hasattr(self, "ent_busca_verba") else ""
        cl = self.cmb_classe.get() if hasattr(self, "cmb_classe") else "Todas"
        out = []
        for v in self.verbas:
            if cl != "Todas" and str(v["CLASSE"]) != cl:
                continue
            if q:
                alvo = f"{v['VERBA']} {v['DESCRICAO']}".lower()
                if q not in alvo:
                    continue
            out.append(v)
        col, rev = self._sort_verbas
        def chave(v):
            if col == "TOTAL":
                return abs(v["TOTAL"])
            if col == "QTD":
                return v["QTD"]
            if col == "DESCRICAO":
                return str(v["DESCRICAO"]).lower()
            return str(v["VERBA"])
        out.sort(key=chave, reverse=rev)
        return out

    def _ordenar_verbas(self, escolha):
        m = {"Codigo": ("VERBA", False), "Maior valor": ("TOTAL", True),
             "Descricao": ("DESCRICAO", False), "Qtd lancamentos": ("QTD", True)}
        self._sort_verbas = m.get(escolha, ("VERBA", False))
        self._view_sort.pop(str(self.tv_verbas), None)
        base = self._head_txt.get(str(self.tv_verbas), {})
        for c, txt in base.items():
            self.tv_verbas.heading(c, text=txt)
        self._render_verbas()

    def _render_verbas(self):
        if not hasattr(self, "tv_verbas"):
            return
        classes = sorted({str(v["CLASSE"]) for v in self.verbas if str(v["CLASSE"])})
        atual = self.cmb_classe.get()
        self.cmb_classe.configure(values=["Todas"] + classes)
        if atual not in ["Todas"] + classes:
            self.cmb_classe.set("Todas")

        for i in self.tv_verbas.get_children():
            self.tv_verbas.delete(i)
        n = 0
        for v in self._verbas_filtradas():
            sel = str(v["VERBA"]) in self.verbas_sel
            tag = ("sel" if n % 2 == 0 else "sel_alt") if sel else ("par" if n % 2 == 0 else "impar")
            n += 1
            self.tv_verbas.insert("", "end", iid=str(v["VERBA"]), tags=(tag,),
                                  values=("[x]" if sel else "[ ]", v["VERBA"], v["DESCRICAO"],
                                          v["CLASSE"], v["QTD"], v["PESSOAS"], v["N_CDC"],
                                          fmt_num(v["TOTAL"])))
        self._sort_view(self.tv_verbas)
        self._kpi_verbas()

    def _kpi_verbas(self):
        total = sum(v["TOTAL"] for v in self.verbas if str(v["VERBA"]) in self.verbas_sel)
        qtd_l = sum(v["QTD"] for v in self.verbas if str(v["VERBA"]) in self.verbas_sel)
        n = len(self.verbas_sel)
        if n:
            self.kv_verba_total.configure(text=fmt_brl(total))
            self.kd_verba_total.configure(text=f"{qtd_l} lancamento(s) em {n} verba(s)")
            self.kv_verba_qtd.configure(text=str(n))
            self.kd_verba_qtd.configure(text=", ".join(sorted(self.verbas_sel)[:6]) +
                                        ("..." if n > 6 else ""))
        else:
            self.kv_verba_total.configure(text="—")
            self.kd_verba_total.configure(text=f"{len(self.verbas)} verbas disponiveis — nenhuma selecionada")
            self.kv_verba_qtd.configure(text="0")
            self.kd_verba_qtd.configure(text="selecione ao menos uma verba")

    def _clique_verba(self, event):
        tv = self.tv_verbas
        if tv.identify_region(event.x, event.y) != "cell":
            return
        iid = tv.identify_row(event.y)
        if not iid:
            return
        if iid in self.verbas_sel:
            self.verbas_sel.discard(iid)
        else:
            self.verbas_sel.add(iid)
        self._render_verbas()

    def _enter_busca_verba(self, event=None):
        f = self._verbas_filtradas()
        if not f:
            return
        cod = str(f[0]["VERBA"])
        if cod in self.verbas_sel:
            self.verbas_sel.discard(cod)
        else:
            self.verbas_sel.add(cod)
        self.ent_busca_verba.delete(0, "end")
        self._render_verbas()

    def _limpar_busca_verba(self):
        self.ent_busca_verba.delete(0, "end")
        self.cmb_classe.set("Todas")
        self._render_verbas()

    def _marcar_verbas(self, marcar):
        for v in self._verbas_filtradas():
            if marcar:
                self.verbas_sel.add(str(v["VERBA"]))
            else:
                self.verbas_sel.discard(str(v["VERBA"]))
        self._render_verbas()

    def _limpar_verbas(self):
        self.verbas_sel = set()
        self._render_verbas()

    def _dclique_verba(self, event):
        iid = self.tv_verbas.identify_row(event.y)
        if not iid:
            return
        itens = [l for l in self._lanc_ativos() if str(l["VERBA"]) == iid]
        desc = itens[0]["DESCRICAO"] if itens else ""
        self._janela_detalhe(f"Verba {iid} — {desc}", itens)

    def _janela_detalhe(self, titulo, itens):
        w = ctk.CTkToplevel(self.app)
        w.title(titulo)
        w.geometry("900x560")
        w.grab_set()
        head = ctk.CTkFrame(w, fg_color=self.AZUL, corner_radius=0, height=58)
        head.pack(fill="x")
        head.pack_propagate(False)
        ctk.CTkLabel(head, text=titulo, font=("Segoe UI", 16, "bold"),
                     text_color="white").pack(side="left", padx=18, pady=14)
        total = sum(i["VALOR"] for i in itens)
        ctk.CTkLabel(head, text=f"{len(itens)} lancamento(s)  |  {fmt_brl(total)}",
                     font=("Segoe UI", 12), text_color="#D6EAF6").pack(side="right", padx=18)
        body = ctk.CTkFrame(w, fg_color=self.CARD)
        body.pack(fill="both", expand=True, padx=12, pady=12)
        cols = [("NOME", "Nome", 280, "w"), ("CPF", "CPF", 130, "center"),
                ("CDC", "CDC", 100, "center"), ("VERBA", "Codigo", 80, "center"),
                ("VALOR", "Valor", 130, "e")]
        wrap, tv = self.tabela(body, cols)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        for n, i in enumerate(sorted(itens, key=lambda x: -abs(x["VALOR"]))):
            tv.insert("", "end", tags=("par" if n % 2 == 0 else "impar",),
                      values=(i["NOME"], i["CPF"], i["CDC"], i["VERBA"], fmt_num(i["VALOR"])))

    # =========================================================
    # ABA 3 — RATEIO
    # =========================================================
    def aba_rateio(self, parent):
        parent = self._rolavel(parent)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(1, weight=1, minsize=300)

        top = self.card(parent)
        top.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.titulo(top, "Rateio por centro de custo",
                    "Valores agregados por CDC. Escolha lancar por VALOR ou por PERCENTUAL — o percentual fecha em 100,00%.")

        bar = ctk.CTkFrame(top, fg_color=self.PAINEL, corner_radius=10,
                           border_width=1, border_color=self.BORDA)
        bar.pack(fill="x", padx=16, pady=(0, 10))

        l1 = ctk.CTkFrame(bar, fg_color="transparent")
        l1.pack(fill="x", padx=12, pady=(10, 6))
        ctk.CTkLabel(l1, text="Lancar como:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(0, 6))
        self.cmb_base_rateio = self.combo(l1, ["VALOR", "PERCENTUAL"],
                                          command=lambda v: self._render_rateio(), width=150)
        self.cmb_base_rateio.set("VALOR")
        self.cmb_base_rateio.pack(side="left", padx=(0, 14))

        ctk.CTkLabel(l1, text="Buscar CDC:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(0, 6))
        self.ent_busca_cdc = self.entry(l1, width=180, placeholder="ex.: 424000")
        self.ent_busca_cdc.pack(side="left", padx=(0, 14))
        self.ent_busca_cdc.bind("<KeyRelease>", lambda e: self._render_rateio())

        ctk.CTkLabel(l1, text="Faixa:", font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(0, 6))
        self.cmb_faixa = self.combo(l1, ["Todas", "Baixo", "Medio", "Alto"],
                                    command=lambda v: self._render_rateio(), width=120)
        self.cmb_faixa.set("Todas")
        self.cmb_faixa.pack(side="left", padx=(0, 14))

        self.botao(l1, "Recalcular", self.gerar_rateio, self.AZUL, self.AZUL_HV, 120).pack(side="left")

        l2 = ctk.CTkFrame(bar, fg_color="transparent")
        l2.pack(fill="x", padx=12, pady=(0, 10))
        self.botao(l2, "Marcar todos", lambda: self._marcar_rateio(True), self.AZUL, self.AZUL_HV, 130).pack(side="left", padx=(0, 8))
        self.botao(l2, "Desmarcar todos", lambda: self._marcar_rateio(False), self.AZUL_CLARO, self.AZUL_CLARO_HV, 145).pack(side="left", padx=(0, 8))
        self.botao(l2, "Inverter selecao", self._inverter_rateio, self.TEAL, self.TEAL_HV, 145).pack(side="left", padx=(0, 8))
        self.botao(l2, "Exportar rateio (CSV)", self.exportar_rateio, self.NEUTRA, self.NEUTRA_HV, 175).pack(side="left")
        self.botao(l2, "Enviar para o SSA  >", self.enviar_para_ssa, self.SUCESSO, self.SUCESSO_HV, 190).pack(side="right")

        kf = ctk.CTkFrame(top, fg_color="transparent")
        kf.pack(fill="x", padx=16, pady=(0, 14))
        kf.grid_columnconfigure(0, weight=2)
        kf.grid_columnconfigure(1, weight=1)
        k1, self.kv_rateio, self.kd_rateio = self.kpi(kf, "TOTAL SELECIONADO A LANCAR NO SSA")
        k1.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._kpi_rateio_frame = k1
        k2, self.kv_pct, self.kd_pct = self.kpi(kf, "SOMA DOS PERCENTUAIS")
        k2.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self._kpi_pct_frame = k2

        body = self.card(parent)
        body.grid(row=1, column=0, sticky="nsew")
        cols = [("SEL", "Sel.", 55, "center"), ("CDC", "CDC", 110, "center"),
                ("PESSOAS", "Pessoas", 90, "center"), ("QTD", "Lancamentos", 100, "center"),
                ("VALOR", "Valor  (duplo clique p/ editar)", 200, "e"),
                ("PCT", "% do total", 100, "e"), ("FAIXA", "Faixa", 90, "center")]
        wrap, self.tv_rateio = self.tabela(body, cols, on_click=self._clique_rateio,
                                           on_dclick=self._editar_valor_rateio, altura=10)
        wrap.pack(fill="both", expand=True, padx=16, pady=(14, 6))
        self.lbl_rateio_rodape = ctk.CTkLabel(
            body, text="Selecione as verbas na aba 2 e clique em Gerar rateio.",
            font=("Segoe UI", 10), text_color=self.TEXTO_SUAVE)
        self.lbl_rateio_rodape.pack(anchor="w", padx=16, pady=(0, 12))

    def gerar_rateio(self):
        if not self.lancamentos:
            messagebox.showwarning("Rateio", "Processe a planilha primeiro (aba 1).")
            return
        if not self.verbas_sel:
            messagebox.showwarning("Rateio", "Selecione ao menos uma verba na aba 2.")
            self.tabs.set(self.T2)
            self._ao_trocar_aba()
            return
        linhas, total = calcular_rateio(self._lanc_ativos(), self.verbas_sel)
        for l in linhas:
            l["SEL"] = True
        self.rateio = linhas
        self.rateio_total = total
        self.faixas = faixas_tercis([l["VALOR"] for l in linhas])
        for l in linhas:
            l["FAIXA"] = classe_faixa(l["VALOR"], self.faixas)
        self._render_rateio()
        self.tabs.set(self.T3)
        self._ao_trocar_aba()
        verbas_txt = ", ".join(sorted(self.verbas_sel))
        self.lbl_rateio_rodape.configure(
            text=f"Empresa: {self.empresa_sel}  |  Verba(s): {verbas_txt}  |  "
                 f"{len(linhas)} CDC(s)  |  total {fmt_brl(total)}")
        self.log(f"Rateio gerado: verbas [{verbas_txt}] -> {len(linhas)} CDC(s), total {fmt_brl(total)}.")

    def _rateio_filtrado(self):
        q = self.ent_busca_cdc.get().strip().lower() if hasattr(self, "ent_busca_cdc") else ""
        fx = self.cmb_faixa.get() if hasattr(self, "cmb_faixa") else "Todas"
        out = []
        for l in self.rateio:
            if q and q not in str(l["CDC"]).lower():
                continue
            if fx != "Todas" and l.get("FAIXA") != fx:
                continue
            out.append(l)
        return out

    def _recalc_pct(self):
        total = round(sum(l["VALOR"] for l in self.rateio if l.get("SEL")), 2)
        self.rateio_total = total
        for l in self.rateio:
            l["PCT"] = round(l["VALOR"] / total * 100.0, 2) if (total and l.get("SEL")) else 0.0
        sel = [l for l in self.rateio if l.get("SEL")]
        if sel and total:
            dif = round(100.0 - round(sum(l["PCT"] for l in sel), 2), 2)
            if abs(dif) >= 0.005:
                alvo = max(sel, key=lambda x: abs(x["VALOR"]))
                alvo["PCT"] = round(alvo["PCT"] + dif, 2)
        self.faixas = faixas_tercis([l["VALOR"] for l in self.rateio])
        for l in self.rateio:
            l["FAIXA"] = classe_faixa(l["VALOR"], self.faixas)

    def _render_rateio(self):
        if not hasattr(self, "tv_rateio"):
            return
        self._recalc_pct()
        for i in self.tv_rateio.get_children():
            self.tv_rateio.delete(i)
        for n, l in enumerate(self._rateio_filtrado()):
            sel = l.get("SEL", False)
            tag = ("sel" if n % 2 == 0 else "sel_alt") if sel else "off"
            self.tv_rateio.insert("", "end", iid=str(l["CDC"]), tags=(tag,),
                                  values=("[x]" if sel else "[ ]", l["CDC"], l["PESSOAS"], l["QTD"],
                                          fmt_num(l["VALOR"]),
                                          (fmt_num(l["PCT"]) + " %") if sel else "—",
                                          l.get("FAIXA", "—")))
        self._sort_view(self.tv_rateio)
        self._kpi_rateio()

    def _kpi_rateio(self):
        sel = [l for l in self.rateio if l.get("SEL")]
        total = sum(l["VALOR"] for l in sel)
        soma_pct = round(sum(l["PCT"] for l in sel), 2)
        modo = self.cmb_base_rateio.get() if hasattr(self, "cmb_base_rateio") else "VALOR"
        if sel:
            self._kpi_rateio_frame.configure(fg_color=self.VERDE_KPI)
            self.kv_rateio.configure(text=fmt_brl(total))
            self.kd_rateio.configure(
                text=f"{len(sel)} CDC(s)  ·  {sum(l['QTD'] for l in sel)} lancamento(s)  ·  lancar por {modo}",
                text_color="#A8D5B5")
            fecha = abs(soma_pct - 100.0) < 0.005
            self._kpi_pct_frame.configure(fg_color=self.VERDE_KPI if fecha else self.ALERTA)
            self.kv_pct.configure(text=f"{fmt_num(soma_pct)} %")
            self.kd_pct.configure(text="fecha em 100,00%" if fecha else "nao fecha — revise",
                                  text_color="#A8D5B5" if fecha else "#F2C9C0")
        else:
            self._kpi_rateio_frame.configure(fg_color=self.AZUL_ESC)
            self.kv_rateio.configure(text="—")
            self.kd_rateio.configure(text=f"{len(self.rateio)} CDC(s) no rateio · nada selecionado",
                                     text_color="#9EC8E8")
            self._kpi_pct_frame.configure(fg_color=self.AZUL_ESC)
            self.kv_pct.configure(text="—")
            self.kd_pct.configure(text="aguardando selecao", text_color="#9EC8E8")

    def _clique_rateio(self, event):
        tv = self.tv_rateio
        if tv.identify_region(event.x, event.y) != "cell":
            return
        col = tv.identify_column(event.x)
        iid = tv.identify_row(event.y)
        if not iid or tv["columns"][int(col.replace("#", "")) - 1] != "SEL":
            return
        for l in self.rateio:
            if str(l["CDC"]) == iid:
                l["SEL"] = not l.get("SEL", False)
                break
        self._render_rateio()

    def _marcar_rateio(self, marcar):
        alvos = {str(l["CDC"]) for l in self._rateio_filtrado()}
        for l in self.rateio:
            if str(l["CDC"]) in alvos:
                l["SEL"] = marcar
        self._render_rateio()

    def _inverter_rateio(self):
        alvos = {str(l["CDC"]) for l in self._rateio_filtrado()}
        for l in self.rateio:
            if str(l["CDC"]) in alvos:
                l["SEL"] = not l.get("SEL", False)
        self._render_rateio()

    def _editar_valor_rateio(self, event):
        tv = self.tv_rateio
        if tv.identify_region(event.x, event.y) != "cell":
            return
        col_id = tv.identify_column(event.x)
        iid = tv.identify_row(event.y)
        if not iid:
            return
        col_idx = int(col_id.replace("#", "")) - 1
        nome_col = tv["columns"][col_idx]
        if nome_col != "VALOR":
            # duplo clique fora da coluna Valor abre o detalhe do CDC
            itens = [l for l in self._lanc_ativos()
                     if str(l["VERBA"]) in self.verbas_sel and (l["CDC"] or "(sem CDC)") == iid]
            self._janela_detalhe(f"CDC {iid} — lancamentos das verbas selecionadas", itens)
            return
        bbox = tv.bbox(iid, column=col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        atual = tv.item(iid, "values")[col_idx]
        ent = tk.Entry(tv, font=("Segoe UI", 11), justify="right", relief="flat", bd=1,
                       highlightthickness=1, highlightbackground=self.AZUL, highlightcolor=self.AZUL)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, atual)
        ent.select_range(0, "end")
        ent.focus_set()

        def confirmar(e=None):
            novo = ent.get().strip()
            ent.destroy()
            v = parse_valor(novo)
            if v is None:
                return
            for l in self.rateio:
                if str(l["CDC"]) == iid:
                    l["VALOR"] = round(v, 2)
                    break
            self._render_rateio()

        def cancelar(e=None):
            ent.destroy()

        ent.bind("<Return>", confirmar)
        ent.bind("<KP_Enter>", confirmar)
        ent.bind("<Escape>", cancelar)
        ent.bind("<FocusOut>", confirmar)

    def exportar_rateio(self):
        if not self.rateio:
            messagebox.showwarning("Exportar", "Nao ha rateio gerado.")
            return
        caminho = filedialog.asksaveasfilename(
            title="Exportar rateio", defaultextension=".csv",
            filetypes=[("CSV", "*.csv")], initialfile="rateio_ssa.csv")
        if not caminho:
            return
        try:
            sel = [l for l in self.rateio if l.get("SEL")]
            df = pd.DataFrame([{"CDC": l["CDC"], "PESSOAS": l["PESSOAS"], "LANCAMENTOS": l["QTD"],
                                "VALOR": l["VALOR"], "PERCENTUAL": l["PCT"], "FAIXA": l.get("FAIXA", "")}
                               for l in sel])
            df.to_csv(caminho, sep=";", index=False, decimal=",", encoding="utf-8-sig")
            self.log(f"Rateio exportado: {caminho}")
            messagebox.showinfo("Exportar", f"Rateio exportado:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def enviar_para_ssa(self):
        sel = [l for l in self.rateio if l.get("SEL")]
        if not sel:
            messagebox.showwarning("SSA", "Selecione ao menos um CDC no rateio.")
            return
        modo = self.cmb_base_rateio.get()
        registros = []
        for l in sel:
            if modo == "PERCENTUAL":
                dado = f"{l['PCT']:.2f}".replace(".", ",")
            else:
                dado = f"{l['VALOR']:.2f}".replace(".", ",")
            obs = f"{l['QTD']} lancamento(s) · {l['PESSOAS']} pessoa(s)"
            registros.append({"CDC": l["CDC"], "DADO": dado, "TIPO": modo,
                              "STATUS": "PENDENTE", "OBS": obs})
        self.df = pd.DataFrame(registros)
        self.cmb_usar.set(modo)
        self.pre_validar_dataframe()
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.tabs.set(self.T4)
        self._ao_trocar_aba()
        total = sum(l["VALOR"] for l in sel)
        self.log(f"Enviado ao SSA: {len(registros)} CDC(s) por {modo} — total {fmt_brl(total)}.")
        messagebox.showinfo("Pronto para o SSA",
                            f"{len(registros)} CDC(s) carregado(s) por {modo}.\n"
                            f"Total: {fmt_brl(total)}\n\n"
                            "Clique no primeiro campo do SSA e pressione F8.\n"
                            "Ou ative o Modo travado para preencher todos automaticamente.")

    # =========================================================
    # PERFIL (mapeamento + verbas favoritas)
    # =========================================================
    def salvar_perfil(self):
        try:
            dados = {"mapa": self.mapa, "verbas": sorted(self.verbas_sel),
                     "empresa": self.empresa_sel,
                     "base_rateio": self.cmb_base_rateio.get() if hasattr(self, "cmb_base_rateio") else "VALOR",
                     "ssa": {"usar": self.cmb_usar.get(), "delay": self.entry_delay.get(),
                             "delay_inicial": self.entry_delay_inicial.get(),
                             "tabs_cdc": self.entry_tabs_cdc.get(),
                             "tabs_final": self.entry_tabs_final.get(),
                             "limpar": self.cmb_limpar.get(),
                             "acao_final": self.cmb_acao_final.get(),
                             "linha_inicial": self.entry_linha_inicial.get(),
                             "qtd": self.entry_qtd.get(),
                             "delay_loop": self.entry_delay_loop.get()
                             if hasattr(self, "entry_delay_loop") else "0,40"}}
            caminho = filedialog.asksaveasfilename(
                title="Salvar perfil", defaultextension=".json",
                filetypes=[("JSON", "*.json")], initialfile="perfil_ssa.json")
            if not caminho:
                return
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=4)
            self.log(f"Perfil salvo: {caminho}")
            messagebox.showinfo("Perfil", "Perfil salvo (mapeamento, verbas e configuracao do SSA).")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def carregar_perfil(self):
        try:
            caminho = filedialog.askopenfilename(title="Carregar perfil", filetypes=[("JSON", "*.json")])
            if not caminho:
                return
            with open(caminho, "r", encoding="utf-8") as f:
                d = json.load(f)
            mapa = d.get("mapa", {})
            if self.df_raw is not None:
                cols = list(self.df_raw.columns)
                for k, cb in self.cmb_map.items():
                    v = mapa.get(k, "")
                    cb.set(v if v in cols else "(nenhuma)")
                    self.mapa[k] = v if v in cols else ""
                self.processar_planilha(silencioso=True)
            else:
                self.mapa = mapa
            emp = d.get("empresa", "(Todas as empresas)")
            if emp in list(self.cmb_empresa.cget("values")):
                self.empresa_sel = emp
                self.cmb_empresa.set(emp)
                self.verbas = resumo_verbas(self._lanc_ativos())
            self.verbas_sel = set(str(v) for v in d.get("verbas", []))
            if hasattr(self, "cmb_base_rateio"):
                self.cmb_base_rateio.set(d.get("base_rateio", "VALOR"))
            s = d.get("ssa", {})
            if s:
                self.cmb_usar.set(s.get("usar", "VALOR"))
                pares = [(self.entry_delay, "delay", "0,15"),
                         (self.entry_delay_inicial, "delay_inicial", "0,15"),
                         (self.entry_tabs_cdc, "tabs_cdc", "3"),
                         (self.entry_tabs_final, "tabs_final", "9"),
                         (self.entry_linha_inicial, "linha_inicial", "1"),
                         (self.entry_qtd, "qtd", "TODOS")]
                if hasattr(self, "entry_delay_loop"):
                    pares.append((self.entry_delay_loop, "delay_loop", "0,80"))
                for ent, key, dflt in pares:
                    ent.delete(0, "end")
                    ent.insert(0, s.get(key, dflt))
                self.cmb_limpar.set(s.get("limpar", "SIM"))
                self.cmb_acao_final.set(s.get("acao_final", "TAB"))
            self._render_verbas()
            self.log(f"Perfil carregado: {caminho}")
            messagebox.showinfo("Perfil", "Perfil aplicado. Confira as verbas na aba 2.")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # =========================================================
    # ABA 4 — EXECUTAR SSA
    # =========================================================
    def aba_ssa(self, parent):
        parent = self._rolavel(parent)
        parent.grid_columnconfigure(0, weight=3)
        parent.grid_columnconfigure(1, weight=2)
        parent.grid_rowconfigure(0, weight=0)
        parent.grid_rowconfigure(1, weight=1, minsize=300)

        # Config
        c = self.card(parent)
        c.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 8))
        self.titulo(c, "Configuracao do preenchimento",
                    "Ajuste os TABs conforme a tela do SSA. Cole os dados manualmente se preferir.")

        grid = ctk.CTkFrame(c, fg_color="transparent")
        grid.pack(fill="x", padx=16, pady=(0, 8))
        for i in range(5):
            grid.grid_columnconfigure(i, weight=1)

        rotulos1 = ["Usar campo", "Navegacao", "Delay entre acoes", "Delay inicial", "Qtd. a executar"]
        for i, t in enumerate(rotulos1):
            ctk.CTkLabel(grid, text=t, font=self.F_LBL, text_color=self.TEXTO
                         ).grid(row=0, column=i, sticky="w", padx=5, pady=(0, 3))
        self.cmb_usar = self.combo(grid, ["VALOR", "PERCENTUAL"])
        self.cmb_usar.grid(row=1, column=0, sticky="ew", padx=5, pady=(0, 8))
        self.cmb_usar.set("VALOR")
        self.cmb_modo = self.combo(grid, ["TAB"])
        self.cmb_modo.grid(row=1, column=1, sticky="ew", padx=5, pady=(0, 8))
        self.cmb_modo.set("TAB")
        self.entry_delay = self.entry(grid)
        self.entry_delay.grid(row=1, column=2, sticky="ew", padx=5, pady=(0, 8))
        self.entry_delay.insert(0, "0,15")
        self.entry_delay_inicial = self.entry(grid)
        self.entry_delay_inicial.grid(row=1, column=3, sticky="ew", padx=5, pady=(0, 8))
        self.entry_delay_inicial.insert(0, "0,15")
        self.entry_qtd = self.entry(grid)
        self.entry_qtd.grid(row=1, column=4, sticky="ew", padx=5, pady=(0, 8))
        self.entry_qtd.insert(0, "TODOS")

        rotulos2 = ["TAB apos CDC", "TAB apos dado", "Limpar campo atual", "Acao final", "Linha inicial"]
        for i, t in enumerate(rotulos2):
            ctk.CTkLabel(grid, text=t, font=self.F_LBL, text_color=self.TEXTO
                         ).grid(row=2, column=i, sticky="w", padx=5, pady=(6, 3))
        self.entry_tabs_cdc = self.entry(grid)
        self.entry_tabs_cdc.grid(row=3, column=0, sticky="ew", padx=5, pady=(0, 8))
        self.entry_tabs_cdc.insert(0, "3")
        self.entry_tabs_final = self.entry(grid)
        self.entry_tabs_final.grid(row=3, column=1, sticky="ew", padx=5, pady=(0, 8))
        self.entry_tabs_final.insert(0, "9")
        self.cmb_limpar = self.combo(grid, ["SIM", "NAO"])
        self.cmb_limpar.grid(row=3, column=2, sticky="ew", padx=5, pady=(0, 8))
        self.cmb_limpar.set("SIM")
        self.cmb_acao_final = self.combo(grid, ["ENTER", "TAB", "NENHUMA"])
        self.cmb_acao_final.grid(row=3, column=3, sticky="ew", padx=5, pady=(0, 8))
        self.cmb_acao_final.set("TAB")
        self.entry_linha_inicial = self.entry(grid)
        self.entry_linha_inicial.grid(row=3, column=4, sticky="ew", padx=5, pady=(0, 8))
        self.entry_linha_inicial.insert(0, "1")

        # Entrada manual (compatibilidade v5)
        ctk.CTkLabel(c, text="Entrada manual (opcional) — CDC e valor separados por TAB ou ponto e virgula",
                     font=self.F_LBL, text_color=self.TEXTO).pack(anchor="w", padx=16, pady=(4, 4))
        self.txt_entrada = ctk.CTkTextbox(c, height=80, font=("Consolas", 12), fg_color="#FBFCFE",
                                          text_color=self.TEXTO, border_width=1, border_color=self.BORDA)
        self.txt_entrada.pack(fill="x", padx=16, pady=(0, 8))

        b = ctk.CTkFrame(c, fg_color="transparent")
        b.pack(fill="x", padx=16, pady=(0, 14))
        self.botao(b, "Ler dados colados", self.ler_dados_colados, self.AZUL_ESC, self.AZUL_ESC_HV, 155).pack(side="left", padx=(0, 6))
        self.botao(b, "Modelo valor", self.inserir_modelo_valor, self.AZUL_CLARO, self.AZUL_CLARO_HV, 125).pack(side="left", padx=(0, 6))
        self.botao(b, "Modelo percentual", self.inserir_modelo_percentual, self.AZUL_CLARO, self.AZUL_CLARO_HV, 155).pack(side="left", padx=(0, 6))
        self.botao(b, "Validar base", self.validar_base_visual, self.TEAL, self.TEAL_HV, 120).pack(side="left", padx=(0, 6))
        self.botao(b, "Limpar base", self.limpar_base, self.ALERTA, self.ALERTA_HV, 120).pack(side="left")

        # Execucao — ocupa toda a coluna direita (rowspan) para o log respirar
        c2 = self.card(parent)
        c2.grid(row=0, column=1, rowspan=2, sticky="nsew", padx=(6, 0), pady=(0, 0))
        self.titulo(c2, "Execucao por hotkey",
                    "1. Clique no primeiro campo do SSA.  2. F8 = 1 linha  ·  Modo travado = loop automatico.  "
                    "F6 pula, F7 pausa, F9 volta.")

        b1 = ctk.CTkFrame(c2, fg_color="transparent")
        b1.pack(fill="x", padx=16, pady=(0, 8))
        self.btn_hotkey = self.botao(b1, "Ativar hotkey global", self.toggle_global,
                                     self.AZUL, self.AZUL_HV, 180)
        self.btn_hotkey.pack(side="left", padx=(0, 6))
        self.botao(b1, "Executar 1 linha", self.executar_proxima_linha_manual,
                   self.SUCESSO, self.SUCESSO_HV, 145).pack(side="left", padx=(0, 6))

        b2 = ctk.CTkFrame(c2, fg_color="transparent")
        b2.pack(fill="x", padx=16, pady=(0, 10))
        self.botao(b2, "Iniciar da linha", self.iniciar_da_linha_configurada, self.AZUL_CLARO, self.AZUL_CLARO_HV, 135).pack(side="left", padx=(0, 6))
        self.botao(b2, "Pular linha", self.pular_linha_manual, self.AMBAR, self.AMBAR_HV, 110).pack(side="left", padx=(0, 6))
        self.botao(b2, "Voltar linha", self.voltar_linha, self.NEUTRA, self.NEUTRA_HV, 115).pack(side="left", padx=(0, 6))
        self.botao(b2, "Marcar erro", self.marcar_erro_manual, self.ALERTA, self.ALERTA_HV, 115).pack(side="left")

        b3 = ctk.CTkFrame(c2, fg_color="transparent")
        b3.pack(fill="x", padx=16, pady=(0, 10))
        self.botao(b3, "Ignorar selecionada", self.marcar_linha_ignorada, self.NEUTRA, self.NEUTRA_HV, 165).pack(side="left", padx=(0, 6))
        self.botao(b3, "Exportar log", self.exportar_log_txt, self.TEAL, self.TEAL_HV, 125).pack(side="left")

        # --- Modo travado (loop automatico dos CDCs) ---
        b_lock = ctk.CTkFrame(c2, fg_color=self.PAINEL, corner_radius=10,
                              border_width=1, border_color=self.BORDA)
        b_lock.pack(fill="x", padx=16, pady=(0, 10))
        top_lock = ctk.CTkFrame(b_lock, fg_color="transparent")
        top_lock.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(top_lock, text="Modo travado (loop automatico dos CDCs)",
                     font=("Segoe UI", 12, "bold"),
                     text_color=self.AZUL_ESC).pack(side="left")
        ctk.CTkLabel(top_lock, text="Delay entre linhas:",
                     font=self.F_LBL, text_color=self.TEXTO).pack(side="left", padx=(16, 4))
        self.entry_delay_loop = self.entry(top_lock, width=80)
        self.entry_delay_loop.insert(0, "0,40")
        self.entry_delay_loop.pack(side="left")

        row_lock = ctk.CTkFrame(b_lock, fg_color="transparent")
        row_lock.pack(fill="x", padx=10, pady=(4, 6))
        self.btn_travar = self.botao(row_lock, "Travar preenchimento",
                                     self.toggle_travamento,
                                     self.AZUL_ESC, self.AZUL_ESC_HV, 185)
        self.btn_travar.pack(side="left", padx=(0, 6))
        self.botao(row_lock, "Pausar / Retomar (F7)",
                   self.toggle_pausa, self.AMBAR, self.AMBAR_HV, 175).pack(side="left", padx=(0, 6))
        self.botao(row_lock, "Reposicionar ancora",
                   self.reposicionar_ancora, self.TEAL, self.TEAL_HV, 165).pack(side="left")

        self.lbl_ancora = ctk.CTkLabel(b_lock, text="Modo travado: desligado.",
                                       font=("Segoe UI", 11), text_color=self.TEXTO_SUAVE)
        self.lbl_ancora.pack(anchor="w", padx=12, pady=(0, 8))

        self.lbl_status = ctk.CTkLabel(c2, text="Status: aguardando dados.",
                                       font=("Segoe UI", 12, "bold"), text_color=self.AZUL_ESC)
        self.lbl_status.pack(anchor="w", padx=16, pady=(0, 4))
        self.lbl_linha = ctk.CTkLabel(c2, text="Linha atual: 0", font=self.F_SUB,
                                      text_color=self.TEXTO_SUAVE)
        self.lbl_linha.pack(anchor="w", padx=16, pady=(0, 2))
        self.lbl_progresso = ctk.CTkLabel(c2, text="Progresso: 0/0",
                                          font=("Segoe UI", 12, "bold"), text_color=self.TEAL)
        self.lbl_progresso.pack(anchor="w", padx=16, pady=(0, 8))

        ctk.CTkLabel(c2, text="Log operacional", font=self.F_LBL,
                     text_color=self.TEXTO).pack(anchor="w", padx=16, pady=(0, 4))
        lf = ctk.CTkFrame(c2, fg_color="transparent")
        lf.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.txt_log = ctk.CTkTextbox(lf, height=140, font=("Consolas", 11), fg_color="#FBFCFE",
                                      text_color=self.TEXTO, border_width=1, border_color=self.BORDA,
                                      corner_radius=8, wrap="word",
                                      scrollbar_button_color="#B9CBD8",
                                      scrollbar_button_hover_color=self.AZUL)
        self.txt_log.pack(fill="both", expand=True)

        # Fila do SSA — coluna esquerda, com altura garantida
        c3 = self.card(parent)
        c3.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(0, 0))
        cab = ctk.CTkFrame(c3, fg_color="transparent")
        cab.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(cab, text="Fila de lancamento no SSA", font=self.F_TIT,
                     text_color=self.AZUL_ESC).pack(side="left")
        self.lbl_resumo = ctk.CTkLabel(cab, text="Total: 0 | Pendentes: 0 | OK: 0 | Erro: 0 | Ignorados: 0",
                                       font=self.F_LBL, text_color=self.TEXTO_SUAVE)
        self.lbl_resumo.pack(side="right")
        cols = [("LINHA", "#", 50, "center"), ("CDC", "CDC", 110, "center"),
                ("DADO", "Dado", 130, "e"), ("TIPO", "Tipo", 110, "center"),
                ("STATUS", "Status", 110, "center"), ("OBS", "Observacao", 260, "w")]
        wrap, self.tree = self.tabela(c3, cols, altura=8)
        wrap.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.tree.tag_configure("ok", background="#E6F4EC", foreground="#155E2E")
        self.tree.tag_configure("pend", background="#FFFFFF")
        self.tree.tag_configure("erro", background="#FBEAE6", foreground="#8C321E")
        self.tree.tag_configure("ign", background="#F5F5F5", foreground="#8A8A8A")

    # =========================================================
    # ABA 5 — DADOS INICIAIS DO CONTAS A PAGAR
    # =========================================================
    def aba_dados_iniciais(self, parent):
        parent = self._rolavel(parent)
        parent.grid_columnconfigure(0, weight=3)
        parent.grid_columnconfigure(1, weight=2)
        parent.grid_rowconfigure(0, weight=1, minsize=520)

        # Relacao de fornecedores
        c1 = self.card(parent)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 0))
        self.titulo(c1, "Relacao SSA Fornecedores",
                    "Carregue o Excel e escolha a linha. A busca considera fornecedor, codigo e tipo de documento.")

        arq = ctk.CTkFrame(c1, fg_color="transparent")
        arq.pack(fill="x", padx=16, pady=(0, 8))
        self.botao(arq, "Carregar Excel", self.carregar_relacao_fornecedores,
                   self.AZUL, self.AZUL_HV, 145).pack(side="left", padx=(0, 8))
        self.lbl_relacao_arquivo = ctk.CTkLabel(
            arq, text="nenhum arquivo carregado", font=self.F_SUB,
            text_color=self.TEXTO_SUAVE)
        self.lbl_relacao_arquivo.pack(side="left", fill="x", expand=True)

        filtros = ctk.CTkFrame(c1, fg_color="transparent")
        filtros.pack(fill="x", padx=16, pady=(0, 8))
        ctk.CTkLabel(filtros, text="Aba do Excel", font=self.F_LBL,
                     text_color=self.TEXTO).pack(side="left", padx=(0, 6))
        self.cmb_relacao_aba = self.combo(filtros, ["(arquivo unico)"],
                                          command=self._ao_mudar_relacao_aba, width=210)
        self.cmb_relacao_aba.set("(arquivo unico)")
        self.cmb_relacao_aba.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(filtros, text="Buscar", font=self.F_LBL,
                     text_color=self.TEXTO).pack(side="left", padx=(0, 6))
        self.entry_busca_relacao = self.entry(
            filtros, placeholder="fornecedor, codigo ou tipo...")
        self.entry_busca_relacao.pack(side="left", fill="x", expand=True)
        self.entry_busca_relacao.bind("<KeyRelease>", lambda e: self._render_relacao_fornecedores())

        cols = [
            ("FORNECEDOR", "Fornecedor", 250, "w"),
            ("CODIGO", "Codigo", 90, "center"),
            ("TIPO3", "Tipo (3)", 80, "center"),
            ("TIPO_DOCUMENTO", "Tipo de documento", 180, "w"),
            ("CODIGO_TIPO", "Codigo do tipo", 115, "center"),
            ("VALOR", "Documento / Valor", 130, "w"),
        ]
        wrap, self.tree_relacao = self.tabela(c1, cols, altura=14)
        wrap.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.tree_relacao.bind("<<TreeviewSelect>>", self._selecionar_relacao_fornecedor)

        # Dados escolhidos e execucao F8
        c2 = self.card(parent)
        c2.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=(0, 0))
        self.titulo(c2, "Dados iniciais do card no SSA",
                    "F8 preenche um campo por vez: Documento, Fornecedor e Tipo.")

        campos = ctk.CTkFrame(c2, fg_color="transparent")
        campos.pack(fill="x", padx=16, pady=(0, 8))
        campos.grid_columnconfigure(0, weight=1)
        campos.grid_columnconfigure(1, weight=1)

        def campo(linha, coluna, titulo, somente_excel=False):
            ctk.CTkLabel(campos, text=titulo, font=self.F_LBL,
                         text_color=self.TEXTO).grid(
                             row=linha * 2, column=coluna, sticky="w", padx=5, pady=(4, 3))
            ent = self.entry(campos)
            ent.grid(row=linha * 2 + 1, column=coluna, sticky="ew", padx=5, pady=(0, 5))
            if somente_excel:
                ent.configure(state="disabled")
            return ent

        self.entry_doc_inicial = campo(0, 0, "Documento — VALOR")
        self.entry_fornecedor_inicial = campo(0, 1, "2-Fornecedor — CODIGO", True)
        self.entry_tipo3_inicial = campo(1, 0, "Tipo de documento — 3 primeiras letras", True)
        self.entry_codigo_tipo_inicial = campo(1, 1, "Tipo no SSA — CODIGO_DO_TIPO_DE_DOCUMENTO", True)

        op = ctk.CTkFrame(c2, fg_color=self.PAINEL, corner_radius=10,
                          border_width=1, border_color=self.BORDA)
        op.pack(fill="x", padx=16, pady=(4, 10))
        lin_op = ctk.CTkFrame(op, fg_color="transparent")
        lin_op.pack(fill="x", padx=10, pady=(8, 4))
        ctk.CTkLabel(lin_op, text="Delay:", font=self.F_LBL,
                     text_color=self.TEXTO).pack(side="left", padx=(0, 4))
        self.entry_delay_dados = self.entry(lin_op, width=75)
        self.entry_delay_dados.insert(0, "0,08")
        self.entry_delay_dados.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(lin_op, text="Acao apos Tipo:", font=self.F_LBL,
                     text_color=self.TEXTO).pack(side="left", padx=(0, 4))
        self.cmb_acao_tipo = self.combo(lin_op, ["TAB", "ENTER", "NENHUMA"], width=115)
        self.cmb_acao_tipo.set("TAB")
        self.cmb_acao_tipo.pack(side="left")

        self.lbl_dados_etapa = ctk.CTkLabel(
            op, text="Campo 1/3 — clique em Documento no SSA e pressione F8.",
            font=("Segoe UI", 12, "bold"), text_color=self.AZUL_ESC,
            justify="left", wraplength=430)
        self.lbl_dados_etapa.pack(anchor="w", padx=10, pady=(4, 8))

        botoes = ctk.CTkFrame(c2, fg_color="transparent")
        botoes.pack(fill="x", padx=16, pady=(0, 8))
        self.btn_hotkey_dados = self.botao(
            botoes, "Ativar hotkey global", self.toggle_global,
            self.AZUL, self.AZUL_HV, 180)
        self.btn_hotkey_dados.pack(side="left", padx=(0, 6))
        self.botao(botoes, "Executar campo", self.executar_dados_iniciais_f8,
                   self.SUCESSO, self.SUCESSO_HV, 140).pack(side="left", padx=(0, 6))

        navegacao = ctk.CTkFrame(c2, fg_color="transparent")
        navegacao.pack(fill="x", padx=16, pady=(0, 8))
        self.botao(navegacao, "Voltar campo", self.voltar_dados_iniciais,
                   self.AMBAR, self.AMBAR_HV, 140).pack(side="left", padx=(0, 6))
        self.botao(navegacao, "Reiniciar", self.reiniciar_dados_iniciais,
                   self.NEUTRA, self.NEUTRA_HV, 105).pack(side="left")

        instrucoes = (
            "1. Selecione o fornecedor/tipo no Excel.\n"
            "2. Clique em Documento no SSA e pressione F8.\n"
            "3. Clique em 2-Fornecedor e pressione F8.\n"
            "4. Abra Parcelas, clique em Tipo e pressione F8.\n"
            "Se perder o foco, use Voltar campo e repita somente o campo anterior."
        )
        ctk.CTkLabel(c2, text=instrucoes, font=self.F_SUB, text_color=self.TEXTO_SUAVE,
                     justify="left", wraplength=450).pack(anchor="w", padx=16, pady=(2, 8))
        self.lbl_dados_status = ctk.CTkLabel(
            c2, text="Aguardando o Excel.", font=self.F_LBL,
            text_color=self.TEXTO_SUAVE, justify="left", wraplength=450)
        self.lbl_dados_status.pack(anchor="w", padx=16, pady=(0, 14))

    def _texto_relacao(self, valor):
        if valor is None or pd.isna(valor):
            return ""
        s = str(valor).strip()
        if s.lower() in ("nan", "none", "nat"):
            return ""
        if re.fullmatch(r"-?\d+\.0", s):
            return s[:-2]
        return s

    def _prefixo_tipo_documento(self, texto):
        limpo = re.sub(r"[^A-Z0-9]", "", sem_acento(texto).upper())
        return limpo[:3]

    def carregar_relacao_fornecedores(self):
        caminho = filedialog.askopenfilename(
            title="Carregar Relacao SSA Fornecedores",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls *.xlsb"),
                       ("Texto", "*.csv *.txt"), ("Todos", "*.*")])
        if not caminho:
            return
        self.relacao_arquivo = caminho
        abas = abas_excel(caminho)
        valores = abas or ["(arquivo unico)"]
        self.cmb_relacao_aba.configure(values=valores)
        self.cmb_relacao_aba.set(valores[0])
        self.lbl_relacao_arquivo.configure(text=Path(caminho).name)
        self._carregar_relacao_aba(valores[0])

    def _ao_mudar_relacao_aba(self, aba):
        if self.relacao_arquivo:
            self._carregar_relacao_aba(aba)

    def _carregar_relacao_aba(self, aba):
        try:
            nome_aba = None if str(aba).startswith("(") else aba
            df, mapa = ler_relacao_fornecedores(self.relacao_arquivo, nome_aba)
            self.relacao_df = df
            self.relacao_mapa = mapa
            self.relacao_registros = []
            for _, row in df.iterrows():
                def valor(chave):
                    coluna = mapa.get(chave, "")
                    return self._texto_relacao(row[coluna]) if coluna else ""
                registro = {
                    "VALOR": valor("VALOR"),
                    "CODIGO": valor("CODIGO"),
                    "TIPO_DE_DOCUMENTO": valor("TIPO_DE_DOCUMENTO"),
                    "CODIGO_DO_TIPO_DE_DOCUMENTO": valor("CODIGO_DO_TIPO_DE_DOCUMENTO"),
                    "FORNECEDOR": valor("FORNECEDOR"),
                }
                if any(registro[k] for k in ("CODIGO", "TIPO_DE_DOCUMENTO",
                                             "CODIGO_DO_TIPO_DE_DOCUMENTO")):
                    self.relacao_registros.append(registro)
            self.relacao_selecionado = None
            self.reiniciar_dados_iniciais(limpar_selecao=True)
            self._render_relacao_fornecedores()
            self.lbl_dados_status.configure(
                text=f"{len(self.relacao_registros)} registros carregados. Selecione uma linha.",
                text_color=self.SUCESSO)
            self.log(f"Relacao SSA carregada: {Path(self.relacao_arquivo).name} "
                     f"({len(self.relacao_registros)} registros).")
        except Exception as e:
            self.relacao_df = None
            self.relacao_registros = []
            self._render_relacao_fornecedores()
            self.lbl_dados_status.configure(text=f"Erro ao carregar: {e}", text_color=self.ALERTA)
            messagebox.showerror("Relacao SSA Fornecedores", str(e))

    def _render_relacao_fornecedores(self):
        if not hasattr(self, "tree_relacao"):
            return
        itens_atuais = self.tree_relacao.get_children()
        if itens_atuais:
            self.tree_relacao.delete(*itens_atuais)
        termo = sem_acento(self.entry_busca_relacao.get()).lower().strip()
        for idx, r in enumerate(self.relacao_registros):
            busca = sem_acento(" ".join(r.values())).lower()
            if termo and termo not in busca:
                continue
            tipo3 = self._prefixo_tipo_documento(r["TIPO_DE_DOCUMENTO"])
            fornecedor = r["FORNECEDOR"] or "(sem nome no Excel)"
            self.tree_relacao.insert("", "end", iid=str(idx), values=(
                fornecedor, r["CODIGO"], tipo3, r["TIPO_DE_DOCUMENTO"],
                r["CODIGO_DO_TIPO_DE_DOCUMENTO"], r["VALOR"]),
                tags=("par" if len(self.tree_relacao.get_children()) % 2 == 0 else "impar",))

    def _definir_entry(self, entry, valor, bloqueado=False):
        entry.configure(state="normal")
        entry.delete(0, "end")
        entry.insert(0, valor)
        if bloqueado:
            entry.configure(state="disabled")

    def _selecionar_relacao_fornecedor(self, event=None):
        selecao = self.tree_relacao.selection()
        if not selecao:
            return
        try:
            idx = int(selecao[0])
            r = self.relacao_registros[idx]
        except Exception:
            return
        self.relacao_selecionado = idx
        self._definir_entry(self.entry_doc_inicial, r["VALOR"], False)
        self._definir_entry(self.entry_fornecedor_inicial, r["CODIGO"], True)
        self._definir_entry(self.entry_tipo3_inicial,
                            self._prefixo_tipo_documento(r["TIPO_DE_DOCUMENTO"]), True)
        self._definir_entry(self.entry_codigo_tipo_inicial,
                            r["CODIGO_DO_TIPO_DE_DOCUMENTO"], True)
        self.reiniciar_dados_iniciais()
        nome = r["FORNECEDOR"] or "fornecedor sem nome"
        complemento = " Documento veio do Excel." if r["VALOR"] else " Informe o Documento antes do F8."
        self.lbl_dados_status.configure(
            text=f"Selecionado: {nome} | codigo {r['CODIGO']}.{complemento}",
            text_color=self.SUCESSO)

    def reiniciar_dados_iniciais(self, limpar_selecao=False):
        self.dados_iniciais_etapa = 1
        self._atualizar_etapa_dados_iniciais()
        if limpar_selecao and hasattr(self, "entry_doc_inicial"):
            self._definir_entry(self.entry_doc_inicial, "", False)
            self._definir_entry(self.entry_fornecedor_inicial, "", True)
            self._definir_entry(self.entry_tipo3_inicial, "", True)
            self._definir_entry(self.entry_codigo_tipo_inicial, "", True)

    def _atualizar_etapa_dados_iniciais(self):
        if not hasattr(self, "lbl_dados_etapa"):
            return
        textos = {
            1: ("Campo 1/3 — clique em Documento no SSA e pressione F8.", self.AZUL_ESC),
            2: ("Campo 2/3 — clique em 2-Fornecedor no SSA e pressione F8.", self.AMBAR),
            3: ("Campo 3/3 — abra Parcelas, clique em Tipo e pressione F8.", self.AMBAR),
            4: ("Concluido — os tres campos foram preenchidos. Use Reiniciar para outro card.",
                self.SUCESSO),
        }
        texto, cor = textos.get(self.dados_iniciais_etapa, textos[1])
        self.lbl_dados_etapa.configure(text=texto, text_color=cor)

    def voltar_dados_iniciais(self):
        """Retorna somente a etapa da nova aba, sem alterar a fila/rateio do SSA."""
        anterior = self.dados_iniciais_etapa
        self.dados_iniciais_etapa = max(1, self.dados_iniciais_etapa - 1)
        self._atualizar_etapa_dados_iniciais()
        nomes = {1: "Documento", 2: "2-Fornecedor", 3: "Tipo"}
        campo = nomes.get(self.dados_iniciais_etapa, "Documento")
        if anterior == 1:
            texto = "Voce ja esta no primeiro campo: Documento."
        else:
            texto = f"Voltou para {campo}. Clique nesse campo no SSA e pressione F8."
        self.lbl_dados_status.configure(text=texto, text_color=self.AMBAR)
        self.log(f"Dados iniciais: voltou para o campo {campo}.")

    def obter_delay_dados(self):
        try:
            return max(0.03, float(str(self.entry_delay_dados.get()).replace(",", ".")))
        except Exception:
            return 0.08

    def executar_dados_iniciais_f8(self):
        if self._sending:
            return
        agora = time.time()
        if agora - self.last_f8_time < self._debounce_secs:
            return
        self.last_f8_time = agora
        if self.relacao_selecionado is None:
            messagebox.showwarning("Dados iniciais", "Carregue o Excel e selecione um fornecedor/tipo.")
            return

        documento = self.entry_doc_inicial.get().strip()
        fornecedor = self.entry_fornecedor_inicial.get().strip()
        codigo_tipo = self.entry_codigo_tipo_inicial.get().strip()
        delay = self.obter_delay_dados()
        self._sending = True
        try:
            if self.dados_iniciais_etapa == 1:
                if not documento:
                    messagebox.showwarning("Documento", "Informe o valor do Documento antes de pressionar F8.")
                    return
                time.sleep(delay)
                self.colar_texto(documento, limpar=True)
                time.sleep(delay)
                self.dados_iniciais_etapa = 2
                self._atualizar_etapa_dados_iniciais()
                self.lbl_dados_status.configure(
                    text=f"Documento {documento} preenchido. Proximo campo: 2-Fornecedor.",
                    text_color=self.SUCESSO)
                self.log(f"Dados iniciais campo 1: Documento {documento}.")
            elif self.dados_iniciais_etapa == 2:
                if not fornecedor:
                    messagebox.showwarning("Fornecedor", "O registro selecionado nao possui CODIGO.")
                    return
                time.sleep(delay)
                self.colar_texto(fornecedor, limpar=True)
                time.sleep(delay)
                self.dados_iniciais_etapa = 3
                self._atualizar_etapa_dados_iniciais()
                self.lbl_dados_status.configure(
                    text=f"Fornecedor {fornecedor} preenchido. Proximo campo: Tipo, na aba Parcelas.",
                    text_color=self.SUCESSO)
                self.log(f"Dados iniciais campo 2: Fornecedor {fornecedor}.")
            elif self.dados_iniciais_etapa == 3:
                if not codigo_tipo:
                    messagebox.showwarning(
                        "Tipo", "O registro selecionado nao possui CODIGO_DO_TIPO_DE_DOCUMENTO.")
                    return
                time.sleep(delay)
                self.colar_texto(codigo_tipo, limpar=True)
                time.sleep(delay)
                acao = self.cmb_acao_tipo.get()
                if acao in ("TAB", "ENTER"):
                    pyautogui.press(acao.lower())
                    time.sleep(delay)
                self.dados_iniciais_etapa = 4
                self._atualizar_etapa_dados_iniciais()
                self.lbl_dados_status.configure(
                    text=f"Tipo {codigo_tipo} preenchido no SSA.", text_color=self.SUCESSO)
                self.log(f"Dados iniciais campo 3: Tipo {codigo_tipo}. Fluxo concluido.")
            else:
                self.lbl_dados_status.configure(
                    text="Fluxo concluido. Use Reiniciar para iniciar outro card ou Voltar campo para repetir o Tipo.",
                    text_color=self.TEAL)
        except Exception as e:
            self.lbl_dados_status.configure(text=f"Erro no F8: {e}", text_color=self.ALERTA)
            self.log(f"Dados iniciais: erro — {e}")
            messagebox.showerror("Dados iniciais", str(e))
        finally:
            self._sending = False

    # =========================================================
    # LOG / STATUS
    # =========================================================
    def log(self, msg):
        linha = f"[{time.strftime('%H:%M:%S')}] {msg}\n"
        if hasattr(self, "txt_log"):
            self.txt_log.insert("end", linha)
            self.txt_log.see("end")
            self.app.update_idletasks()
        else:
            print(linha, end="")

    def set_status(self, msg, cor=None):
        self.lbl_status.configure(text=f"Status: {msg}", text_color=cor or self.AZUL_ESC)
        self.app.update_idletasks()

    # =========================================================
    # HOTKEYS
    # =========================================================
    def _bind_local_keys(self):
        self.app.bind("<F8>", self._on_f8_local)
        self.app.bind("<F6>", lambda e: self.pular_linha_manual())
        self.app.bind("<F7>", lambda e: self.toggle_pausa())
        self.app.bind("<F9>", lambda e: self.voltar_linha())
        self._local_f8_bound = True

    def _unbind_local_f8(self):
        if self._local_f8_bound:
            try:
                self.app.unbind("<F8>")
            except Exception:
                pass
            self._local_f8_bound = False

    def _on_f8_local(self, event):
        st = getattr(event, "state", 0)
        if st & 0x0001 or st & 0x0004 or st & 0x0008:
            return
        self._acao_f8()

    def _acao_f8(self):
        """Roteamento do F8:
        - se aguardando ancora   -> captura pyautogui.position() e inicia o loop
        - se modo travado ligado -> ignora (o loop cuida)
        - se aba Dados iniciais  -> executa Documento, Fornecedor ou Tipo, campo a campo
        - caso contrario         -> executa 1 linha (comportamento individual)
        """
        if self._awaiting_anchor:
            try:
                x, y = pyautogui.position()
                self.anchor_xy = (int(x), int(y))
                self._awaiting_anchor = False
                self._atualizar_ancora_ui()
                self.log(f"Ancora gravada em ({x}, {y}).")
                self.set_status("ancora ok — loop rodando", self.SUCESSO)
                self._iniciar_thread_travamento()
            except Exception as e:
                self.log(f"Falha ao capturar ancora: {e}")
                self._awaiting_anchor = False
                self._atualizar_ancora_ui()
            return
        if self._locked:
            return  # loop no controle
        if hasattr(self, "T5") and self.tabs.get() == self.T5:
            self.executar_dados_iniciais_f8()
            return
        self.executar_proxima_linha_manual()

    def _sincronizar_botoes_hotkey(self):
        texto = "Desativar hotkey global" if self.global_on else "Ativar hotkey global"
        cor = self.NEUTRA if self.global_on else self.AZUL
        hover = self.NEUTRA_HV if self.global_on else self.AZUL_HV
        for nome in ("btn_hotkey", "btn_hotkey_dados"):
            botao = getattr(self, nome, None)
            if botao is not None:
                botao.configure(text=texto, fg_color=cor, hover_color=hover)

    def toggle_global(self):
        if not self.global_on:
            try:
                self._unbind_local_f8()
                if HAS_PYNPUT and pk is not None:
                    mapping = {
                        "<f8>": lambda: self.app.after(0, self._acao_f8),
                        "<f6>": lambda: self.app.after(0, self.pular_linha_manual),
                        "<f7>": lambda: self.app.after(0, self.toggle_pausa),
                        "<f9>": lambda: self.app.after(0, self.voltar_linha),
                    }
                    self.hk_listener = pk.GlobalHotKeys(mapping)
                    self.hk_listener.start()
                    self.global_on = True
                    self._sincronizar_botoes_hotkey()
                    self.lbl_hotkey_mode.configure(text="Hotkey: Global (pynput)")
                    self.log("Hotkey global ativada (pynput): F8 preenche/ancora, F6 pula, F7 pausa, F9 volta.")
                    self.set_status("hotkey global ativada", self.SUCESSO)
                    return
                if HAS_KEYBOARD_PACKAGE and kb_global is not None:
                    kb_global.add_hotkey("f8", lambda: self.app.after(0, self._acao_f8))
                    kb_global.add_hotkey("f6", lambda: self.app.after(0, self.pular_linha_manual))
                    kb_global.add_hotkey("f7", lambda: self.app.after(0, self.toggle_pausa))
                    kb_global.add_hotkey("f9", lambda: self.app.after(0, self.voltar_linha))
                    self.global_on = True
                    self._sincronizar_botoes_hotkey()
                    self.lbl_hotkey_mode.configure(text="Hotkey: Global (keyboard)")
                    self.log("Hotkey global ativada (keyboard): F8, F6, F7, F9.")
                    self.set_status("hotkey global ativada", self.SUCESSO)
                    return
                detalhes = []
                if PYNPUT_IMPORT_ERROR:
                    detalhes.append(f"pynput: {PYNPUT_IMPORT_ERROR}")
                if KEYBOARD_IMPORT_ERROR:
                    detalhes.append(f"keyboard: {KEYBOARD_IMPORT_ERROR}")
                messagebox.showerror("Hotkey global",
                                     "Nao foi possivel ativar F8 global. Instale 'pynput' ou 'keyboard'.\n\n"
                                     + ("\n".join(detalhes) or "Sem detalhe tecnico."))
                self.app.bind("<F8>", self._on_f8_local)
                self._local_f8_bound = True
            except Exception as e:
                self.app.bind("<F8>", self._on_f8_local)
                self._local_f8_bound = True
                messagebox.showerror("Hotkey global", f"Falha ao ativar.\n\n{e}")
        else:
            try:
                if self.hk_listener:
                    self.hk_listener.stop()
            except Exception:
                pass
            try:
                if HAS_KEYBOARD_PACKAGE and kb_global is not None:
                    kb_global.unhook_all_hotkeys()
            except Exception:
                pass
            self.hk_listener = None
            self.global_on = False
            self.app.bind("<F8>", self._on_f8_local)
            self._local_f8_bound = True
            self._sincronizar_botoes_hotkey()
            self.lbl_hotkey_mode.configure(text="Hotkey: Local")
            self.log("Hotkey global desativada.")
            self.set_status("hotkey global desativada", self.TEAL)

    # =========================================================
    # ENTRADA MANUAL / BASE SSA
    # =========================================================
    def inserir_modelo_valor(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert("1.0", "CDC\tVALOR\n300000\t843,60\n301000\t3532,00\n303000\t2000,00")
        self.cmb_usar.set("VALOR")

    def inserir_modelo_percentual(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert("1.0", "CDC\tPERCENTUAL\n300000\t25\n301000\t35\n303000\t40")
        self.cmb_usar.set("PERCENTUAL")

    def limpar_base(self):
        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Fila do SSA limpa.")
        self.set_status("fila limpa", self.TEAL)

    def detectar_cabecalho(self, p):
        a, b = str(p[0]).strip().lower(), str(p[1]).strip().lower()
        pal = ["cdc", "valor", "percentual", "percent", "%", "dado"]
        return any(x in a for x in pal) or any(x in b for x in pal)

    def separar_linha(self, linha):
        linha = linha.strip()
        if not linha:
            return None
        for sep in ["\t", ";"]:
            if sep in linha:
                partes = [p.strip() for p in linha.split(sep) if p.strip()]
                if len(partes) >= 2:
                    return partes[0], partes[1]
        partes = [p.strip() for p in re.split(r"\s{2,}", linha) if p.strip()]
        if len(partes) >= 2:
            return partes[0], partes[1]
        return None

    def ler_dados_colados(self):
        texto = self.txt_entrada.get("1.0", "end").strip()
        if not texto:
            messagebox.showwarning("Atencao", "Cole os dados antes de ler.")
            return
        linhas = [l for l in texto.splitlines() if l.strip()]
        primeira = self.separar_linha(linhas[0]) if linhas else None
        inicio = 1 if primeira and self.detectar_cabecalho(primeira) else 0
        tipo = self.cmb_usar.get()
        regs = []
        for i in range(inicio, len(linhas)):
            r = self.separar_linha(linhas[i])
            if not r:
                self.log(f"Linha ignorada: {linhas[i]}")
                continue
            regs.append({"CDC": r[0], "DADO": r[1], "TIPO": tipo, "STATUS": "PENDENTE", "OBS": ""})
        if not regs:
            messagebox.showwarning("Atencao", "Nenhum dado valido identificado.")
            return
        self.df = pd.DataFrame(regs)
        self.pre_validar_dataframe()
        self.iniciar_da_linha_configurada(silencioso=True)
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log(f"{len(self.df)} linha(s) carregada(s) manualmente.")
        self.set_status(f"{len(self.df)} linha(s) carregada(s)", self.SUCESSO)

    def pre_validar_dataframe(self):
        if self.df.empty:
            return
        for idx, row in self.df.iterrows():
            if not str(row["CDC"]).strip():
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "CDC vazio"
            elif not str(row["DADO"]).strip():
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "Dado vazio"
            elif self.df.at[idx, "STATUS"] not in ["OK", "IGNORADO"]:
                self.df.at[idx, "STATUS"] = "PENDENTE"

    def validar_base_visual(self):
        if self.df.empty:
            messagebox.showwarning("Validacao", "Nao ha fila carregada.")
            return
        self.pre_validar_dataframe()
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Fila validada.")
        self.set_status("fila validada", self.TEAL)

    def _tag_status(self, status):
        return {"OK": "ok", "ERRO": "erro", "IGNORADO": "ign"}.get(status, "pend")

    def renderizar_base(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.df.empty:
            return
        for idx, row in self.df.iterrows():
            self.tree.insert("", "end", iid=str(idx), tags=(self._tag_status(row["STATUS"]),),
                             values=(idx + 1, row["CDC"], row["DADO"], row["TIPO"],
                                     row["STATUS"], row["OBS"]))
        self._sort_view(self.tree)

    def atualizar_status_linha(self, idx, status, obs=None):
        if idx < 0 or idx >= len(self.df):
            return
        self.df.at[idx, "STATUS"] = status
        if obs is not None:
            self.df.at[idx, "OBS"] = obs
        self.tree.item(str(idx), tags=(self._tag_status(status),),
                       values=(idx + 1, self.df.at[idx, "CDC"], self.df.at[idx, "DADO"],
                               self.df.at[idx, "TIPO"], self.df.at[idx, "STATUS"],
                               self.df.at[idx, "OBS"]))
        self.atualizar_resumo()

    def atualizar_resumo(self):
        if self.df.empty:
            self.lbl_resumo.configure(text="Total: 0 | Pendentes: 0 | OK: 0 | Erro: 0 | Ignorados: 0")
            self.lbl_progresso.configure(text="Progresso: 0/0")
            return
        total = len(self.df)
        pend = int((self.df["STATUS"] == "PENDENTE").sum())
        ok = int((self.df["STATUS"] == "OK").sum())
        erro = int((self.df["STATUS"] == "ERRO").sum())
        ign = int((self.df["STATUS"] == "IGNORADO").sum())
        self.lbl_resumo.configure(
            text=f"Total: {total} | Pendentes: {pend} | OK: {ok} | Erro: {erro} | Ignorados: {ign}")
        self.lbl_progresso.configure(text=f"Progresso: {ok + ign}/{total}")

    def obter_proxima_linha_pendente(self):
        if self.df.empty:
            return None
        p = self.df.index[self.df["STATUS"].isin(["PENDENTE", "ERRO"])]
        return int(p[0]) if len(p) > 0 else None

    def atualizar_linha_atual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.lbl_linha.configure(text="Linha atual: concluido")
            return
        self.lbl_linha.configure(text=f"Linha atual: {idx + 1} de {len(self.df)}")
        try:
            self.tree.selection_set(str(idx))
            self.tree.focus(str(idx))
            self.tree.see(str(idx))
        except Exception:
            pass

    # =========================================================
    # PARAMETROS
    # =========================================================
    def obter_delay(self):
        try:
            return float(str(self.entry_delay.get()).replace(",", "."))
        except Exception:
            return 0.10

    def obter_delay_inicial(self):
        try:
            return float(str(self.entry_delay_inicial.get()).replace(",", "."))
        except Exception:
            return 0.15

    def obter_tabs_cdc(self):
        try:
            return max(0, int(float(str(self.entry_tabs_cdc.get()).replace(",", "."))))
        except Exception:
            return 1

    def obter_tabs_final(self):
        try:
            return max(0, int(float(str(self.entry_tabs_final.get()).replace(",", "."))))
        except Exception:
            return 0

    def obter_linha_inicial(self):
        try:
            linha = int(float(str(self.entry_linha_inicial.get()).replace(",", ".")))
            return max(1, min(linha, len(self.df))) if not self.df.empty else 1
        except Exception:
            return 1

    # =========================================================
    # EXECUCAO SSA
    # =========================================================
    def posicionar_cursor_cdc_mascarado(self, delay=0.08):
        """Posiciona o cursor antes da mascara fixa do CDC sem tentar apaga-la.

        O ponto exibido pelo SSA pertence a mascara do controle, portanto nao
        deve ser validado pelo clipboard nem removido. HOME independe do local
        exato do clique dentro do campo e leva o cursor para o primeiro espaco
        editavel, preservando o comportamento do codigo original que funciona.
        """
        pyautogui.press("home")
        time.sleep(max(0.05, delay))

    def colar_texto(self, texto, limpar=False):
        pyperclip.copy(str(texto))
        if limpar:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.08)
        pyautogui.hotkey("ctrl", "v")

    def digitar_texto(self, texto, intervalo=0.02):
        pyautogui.write(str(texto), interval=intervalo)

    def limpar_campo_atual(self, delay=0.08):
        pyautogui.hotkey("ctrl", "a")
        time.sleep(delay)
        pyautogui.press("backspace")
        time.sleep(delay)

    def preparar_dado(self, dado, usar):
        v = str(dado).strip()
        return v.replace(" ", "") if usar == "VALOR" else v.replace("%", "").strip()

    def executar_tabs(self, qtd, delay):
        for _ in range(qtd):
            pyautogui.press("tab")
            time.sleep(delay)

    def executar_acao_final(self, delay):
        acao = self.cmb_acao_final.get()
        if acao in ("ENTER", "TAB"):
            pyautogui.press(acao.lower())
            time.sleep(delay)

    def iniciar_da_linha_configurada(self, silencioso=False):
        if self.df.empty:
            if not silencioso:
                messagebox.showwarning("Inicio", "Carregue a fila primeiro.")
            return
        linha = self.obter_linha_inicial() - 1
        for idx in range(len(self.df)):
            if idx < linha and self.df.at[idx, "STATUS"] in ["PENDENTE", "ERRO"]:
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "Ignorado por linha inicial"
            elif idx >= linha and self.df.at[idx, "STATUS"] not in ["OK", "IGNORADO"]:
                self.df.at[idx, "STATUS"] = "PENDENTE"
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log(f"Posicionado na linha {linha + 1}.")

    def executar_linha(self, idx, bypass_debounce=False):
        if idx is None or idx >= len(self.df) or self._sending:
            return False
        if not bypass_debounce:
            agora = time.time()
            if agora - self.last_f8_time < self._debounce_secs:
                return False
            self.last_f8_time = agora
        self._sending = True
        try:
            cdc = str(self.df.at[idx, "CDC"]).strip()
            dado = str(self.df.at[idx, "DADO"]).strip()
            usar = self.cmb_usar.get()
            delay = self.obter_delay()
            if not cdc:
                self.atualizar_status_linha(idx, "IGNORADO", "CDC vazio")
                self.atualizar_linha_atual()
                return True
            if not dado:
                self.atualizar_status_linha(idx, "IGNORADO", "Dado vazio")
                self.atualizar_linha_atual()
                return True
            di = self.obter_delay_inicial()
            if di > 0:
                time.sleep(di)
            # O CDC usa mascara fixa no SSA. Nao selecionar nem apagar o ponto:
            # apenas garantir o cursor no inicio e colar como no fluxo original.
            self.posicionar_cursor_cdc_mascarado(delay=max(0.08, delay))
            self.colar_texto(cdc)
            time.sleep(delay)
            self.executar_tabs(self.obter_tabs_cdc(), delay)
            if self.cmb_limpar.get() == "SIM":
                self.limpar_campo_atual(delay=max(0.05, delay))
            self.digitar_texto(self.preparar_dado(dado, usar), intervalo=max(0.01, delay / 5))
            time.sleep(delay)
            self.executar_tabs(self.obter_tabs_final(), delay)
            self.executar_acao_final(delay)
            self.atualizar_status_linha(idx, "OK", "Preenchido com F8")
            self.log(f"Linha {idx + 1}: OK ({cdc} = {dado}).")
            self.set_status(f"linha {idx + 1} preenchida", self.SUCESSO)
            self.atualizar_linha_atual()
            return True
        except Exception as e:
            self.atualizar_status_linha(idx, "ERRO", str(e))
            self.log(f"Linha {idx + 1}: erro — {e}")
            self.set_status(f"erro na linha {idx + 1}", self.ALERTA)
            self.atualizar_linha_atual()
            return False
        finally:
            time.sleep(0.02)
            self._sending = False

    def executar_proxima_linha_manual(self):
        if self.df.empty:
            messagebox.showwarning("Execucao", "Nao ha fila carregada. Gere o rateio ou cole os dados.")
            return
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.set_status("nao ha linhas pendentes", self.TEAL)
            return
        self.executar_linha(idx)

    def pular_linha_manual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            return
        self.atualizar_status_linha(idx, "IGNORADO", "Pulado manualmente")
        self.log(f"Linha {idx + 1} pulada.")
        self.set_status(f"linha {idx + 1} pulada", self.AMBAR)
        self.atualizar_linha_atual()

    def marcar_erro_manual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            return
        self.atualizar_status_linha(idx, "ERRO", "Marcado manualmente")
        self.set_status(f"linha {idx + 1} com erro", self.ALERTA)
        self.atualizar_linha_atual()

    def voltar_linha(self):
        if self.df.empty:
            return
        linhas = self.df.index[self.df["STATUS"].isin(["OK", "IGNORADO", "ERRO"])].tolist()
        if not linhas:
            self.set_status("nenhuma linha concluida para retornar", self.TEAL)
            return
        idx = linhas[-1]
        self.atualizar_status_linha(idx, "PENDENTE", "Retornada manualmente")
        self.log(f"Linha {idx + 1} retornada para PENDENTE.")
        self.atualizar_linha_atual()

    def marcar_linha_ignorada(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Ignorar", "Selecione uma linha na fila.")
            return
        idx = int(sel[0])
        self.atualizar_status_linha(idx, "IGNORADO", "Ignorada manualmente")
        self.atualizar_linha_atual()

    def exportar_log_txt(self):
        conteudo = self.txt_log.get("1.0", "end").strip()
        if not conteudo:
            messagebox.showwarning("Exportar", "O log esta vazio.")
            return
        caminho = filedialog.asksaveasfilename(title="Salvar log", defaultextension=".txt",
                                               filetypes=[("Texto", "*.txt")],
                                               initialfile="log_ssa.txt")
        if not caminho:
            return
        with open(caminho, "w", encoding="utf-8") as f:
            f.write(conteudo)
        self.log(f"Log exportado: {caminho}")

    # =========================================================
    # MODO TRAVADO — LOOP AUTOMATICO DOS CDCs
    # =========================================================
    def _delay_loop(self):
        try:
            return max(0.20, float(str(self.entry_delay_loop.get()).replace(",", ".")))
        except Exception:
            return 0.40

    def _aplicar_tempos_modo_travado(self):
        """Aplica o preset de velocidade seguro sempre que o travamento e ativado."""
        preset = (
            (self.entry_delay, "0,08"),
            (self.entry_delay_inicial, "0,08"),
            (self.entry_delay_loop, "0,40"),
        )
        for campo, valor in preset:
            campo.delete(0, "end")
            campo.insert(0, valor)
        self.log("Preset do modo travado aplicado: acoes 0,08 s | inicial 0,08 s | linhas 0,40 s.")

    def toggle_travamento(self):
        """Liga/desliga o modo travado. Ao ligar, aguarda F8 no primeiro campo do SSA
        para gravar a ancora (coordenada XY de retorno antes de cada linha)."""
        if self._locked:
            self._parar_travamento("modo travado desativado manualmente")
            return
        if self.df.empty:
            messagebox.showwarning("Modo travado", "Carregue a fila primeiro.")
            return
        if self.obter_proxima_linha_pendente() is None:
            messagebox.showinfo("Modo travado", "Nao ha linhas pendentes.")
            return
        self._aplicar_tempos_modo_travado()
        self._locked = True
        self._paused = False
        self._awaiting_anchor = True
        self._falhas_seguidas = 0
        self.btn_travar.configure(text="Desativar travamento",
                                  fg_color=self.ALERTA, hover_color=self.ALERTA_HV)
        self._atualizar_ancora_ui()
        self.log("Modo travado ATIVADO. Clique no primeiro campo do SSA e pressione F8 (grava a ancora).")
        self.set_status("clique no primeiro campo do SSA e pressione F8 (grava ancora)", self.AMBAR)

    def _iniciar_thread_travamento(self):
        if self._lock_thread and self._lock_thread.is_alive():
            return
        self._lock_thread = threading.Thread(target=self._loop_travado, daemon=True)
        self._lock_thread.start()

    def _loop_travado(self):
        """Roda em thread separada. Reposiciona por clique na ancora antes de cada linha.
        Auto-pausa apos 2 falhas seguidas. Respeita pausa e nova captura de ancora."""
        while self._locked:
            if self._paused or self._awaiting_anchor:
                time.sleep(0.20)
                continue
            idx = self.obter_proxima_linha_pendente()
            if idx is None:
                self.app.after(0, lambda: self._parar_travamento("fila concluida"))
                break
            # Reposicionamento por ancora — garantia de posicao antes de cada linha
            if self.anchor_xy:
                try:
                    pyautogui.click(self.anchor_xy[0], self.anchor_xy[1])
                    time.sleep(self.obter_delay_inicial())
                except Exception as e:
                    self.app.after(0, lambda err=e: self.log(f"Falha ao reposicionar ancora: {err}"))
                    self._paused = True
                    self.app.after(0, self._atualizar_ancora_ui)
                    continue
            ok = self.executar_linha(idx, bypass_debounce=True)
            if not ok:
                self._falhas_seguidas += 1
                if self._falhas_seguidas >= 2:
                    self._paused = True
                    self.app.after(0, self._atualizar_ancora_ui)
                    self.app.after(0, lambda: self.set_status(
                        "pausa automatica por 2 falhas seguidas — reposicione a ancora e retome",
                        self.ALERTA))
                    self.app.after(0, lambda: self.log(
                        "Loop pausado automaticamente (falhas seguidas)."))
            else:
                self._falhas_seguidas = 0
            time.sleep(self._delay_loop())

    def toggle_pausa(self):
        """Pausa ou retoma o loop. Sem efeito quando o modo travado esta desligado."""
        if not self._locked:
            self.set_status("modo travado nao esta ativo", self.TEXTO_SUAVE)
            return
        self._paused = not self._paused
        self._atualizar_ancora_ui()
        if self._paused:
            self.log("Loop PAUSADO.")
            self.set_status("loop pausado — reposicione a ancora se preciso e retome com F7",
                            self.AMBAR)
        else:
            self._falhas_seguidas = 0
            self.log("Loop RETOMADO.")
            self.set_status("loop retomado", self.SUCESSO)

    def reposicionar_ancora(self):
        """Pausa o loop e aguarda um novo F8 para regravar a ancora. Nao apaga o progresso."""
        if not self._locked:
            messagebox.showinfo("Reposicionar ancora",
                                "Ative o modo travado antes de reposicionar a ancora.")
            return
        self._paused = True
        self._awaiting_anchor = True
        self._atualizar_ancora_ui()
        self.log("Aguardando nova ancora. Clique no primeiro campo do SSA e pressione F8.")
        self.set_status("clique no primeiro campo do SSA e pressione F8 (nova ancora)",
                        self.AMBAR)

    def _parar_travamento(self, motivo=""):
        self._locked = False
        self._paused = False
        self._awaiting_anchor = False
        self._falhas_seguidas = 0
        if hasattr(self, "btn_travar"):
            self.btn_travar.configure(text="Travar preenchimento",
                                      fg_color=self.AZUL_ESC,
                                      hover_color=self.AZUL_ESC_HV)
        self._atualizar_ancora_ui()
        if motivo:
            self.log(f"Modo travado desativado ({motivo}).")
            self.set_status(motivo, self.TEAL)

    def _atualizar_ancora_ui(self):
        if not hasattr(self, "lbl_ancora"):
            return
        if not self._locked:
            self.lbl_ancora.configure(
                text="Modo travado: desligado.",
                text_color=self.TEXTO_SUAVE)
            return
        if self._awaiting_anchor:
            self.lbl_ancora.configure(
                text="Modo travado: AGUARDANDO ANCORA (clique no primeiro campo do SSA e pressione F8).",
                text_color=self.AMBAR)
            return
        xy = f"({self.anchor_xy[0]}, {self.anchor_xy[1]})" if self.anchor_xy else "—"
        estado = "PAUSADO" if self._paused else "RODANDO"
        self.lbl_ancora.configure(
            text=f"Modo travado: {estado}  ·  ancora {xy}",
            text_color=self.ALERTA if self._paused else self.SUCESSO)

    # =========================================================
    # FECHAMENTO
    # =========================================================
    def fechar_app(self):
        self._locked = False
        self._paused = False
        self._awaiting_anchor = False
        try:
            if self.hk_listener:
                self.hk_listener.stop()
        except Exception:
            pass
        try:
            if HAS_KEYBOARD_PACKAGE and kb_global is not None:
                kb_global.unhook_all_hotkeys()
        except Exception:
            pass
        self.app.destroy()

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    PreenchedorSSA().run()

# =========================================================
# BUILD (opcional — apps internos Sonova rodam via python.exe direto)
# cd "C:\_RPA"
# Remove-Item -Recurse -Force "C:\_RPA\SSA_v7_1" -ErrorAction SilentlyContinue
# & python.exe -m cx_Freeze PreenchedorSSA_v7.py `
#   --base-name Win32GUI `
#   --target-dir "C:\_RPA\SSA_v7_1" `
#   --target-name "PreenchedorSSA.exe" `
#   --include-modules=customtkinter,darkdetect,pandas,openpyxl,pyautogui,pyperclip,pynput,PIL
# =========================================================

"""
DivergênciaPonto Mailer | v1.0
Lê arquivo TXT de divergência de ponto, agrupa por gestor (chefia),
e cria rascunhos no Outlook com a tabela de divergências no corpo do e-mail.

Fluxo:
  1. Selecionar o TXT exportado do sistema de ponto
  2. Colar a base de gestores (Matrícula Chefia;Nome;Email) na interface
  3. Clicar em "Criar Rascunhos" → gera 1 rascunho por gestor no Outlook

Baseado no CustomerThinker v3.4 — sem proteção de PDF, sem anexo.
"""
from __future__ import annotations

import csv
import os
import re
import time
import traceback
import threading
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

try:
    import customtkinter as ctk
    CTK_AVAILABLE = True
except Exception:
    ctk = None
    CTK_AVAILABLE = False

import pandas as pd
from tkinter import filedialog, messagebox

# ─── Constantes ────────────────────────────────────────────────────────────────
APP_NAME = "DivergênciaPonto Mailer"
VERSAO = "1.0"
LOG_NAME = "log_envio_divergencia_ponto.csv"
FOOTER_UI = "Anderson Marinho | Igarapé Digital"

# Colunas esperadas no TXT (pipe-delimited, encoding latin1/iso-8859-1)
COLUNAS_TXT = [
    "Data Inicial", "Data Final", "Cod Estab", "Nome Estab",
    "Cod CR", "Nome CR", "Matricula Chefia", "Nome Chefia",
    "Matricula", "Nome", "Funcao", "Cracha",
    "Cod Escala", "Cod Turma", "Horario", "Descricao Horario",
    "Data", "DIA", "Batida", "Situacao", "Descricao", "Quantidade",
]

SUBJECT_TEMPLATE = "Divergências de Ponto — Período {periodo} — {nome_chefia}"
BODY_HEADER_TEMPLATE = (
    "Olá {nome_chefia},\n\n"
    "Seguem abaixo as divergências de ponto da sua equipe "
    "referentes ao período de {periodo}.\n\n"
    "Por favor, preencha a coluna \"Observação do Gestor\" com a justificativa "
    "de cada ocorrência e responda este e-mail ao RH.\n"
    "A partir desse mês, estamos concentrando as correções por meio digital, suprimindo o papel.\n"
    "Solicitamos que as tratativas sejam realizadas até hoje __/__/____ às __h.\n\n"
)
BODY_FOOTER_TEMPLATE = (
    "\n\nCaso tenha dúvidas, entre em contato com o RH.\n\n"
    "Atenciosamente,\n"
    "Recursos Humanos\n"
)


# ─── Dataclasses ───────────────────────────────────────────────────────────────
@dataclass
class Gestor:
    matricula_chefia: str
    nome: str
    email: str


@dataclass
class DivergenciaLinha:
    nome_cr: str
    matricula: str
    nome: str
    funcao: str
    horario: str
    data: str
    dia: str
    batida: str
    situacao: str
    descricao: str
    quantidade: str


@dataclass
class ResultadoProcesso:
    matricula_chefia: str
    nome_chefia: str
    email: str
    qtd_colaboradores: int
    qtd_divergencias: int
    status: str  # OK, PULADO, FALHA
    detalhe: str


# ─── Funções utilitárias ──────────────────────────────────────────────────────
def normalizar_texto(valor: str) -> str:
    valor = str(valor or "").strip().upper()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"[^A-Z0-9 ]", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def _normalizar_cabecalho(cab: str) -> str:
    return normalizar_texto(cab).replace(" ", "")


def validar_email_basico(email: str) -> bool:
    if not email:
        return False
    e = email.strip()
    if " " in e or e.count("@") != 1:
        return False
    local, dom = e.split("@")
    if not local or not dom or "." not in dom:
        return False
    return True


def detectar_encoding(caminho: Path) -> str:
    """Tenta detectar encoding do arquivo."""
    for enc in ["utf-8-sig", "utf-8", "latin-1", "cp1252"]:
        try:
            with open(caminho, "r", encoding=enc) as f:
                f.read(4096)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return "latin-1"


# ─── Leitura do TXT de ponto ──────────────────────────────────────────────────
def ler_txt_ponto(caminho: Path) -> pd.DataFrame:
    """Lê o TXT pipe-delimited e retorna DataFrame."""
    enc = detectar_encoding(caminho)
    df = pd.read_csv(
        caminho,
        sep="|",
        encoding=enc,
        dtype=str,
        keep_default_na=False,
        on_bad_lines="skip",
    )
    # Limpa espaços dos cabeçalhos
    df.columns = [c.strip() for c in df.columns]
    # Limpa espaços dos valores
    for col in df.columns:
        df[col] = df[col].str.strip()
    return df


def agrupar_por_chefia(df: pd.DataFrame) -> Dict[str, List[DivergenciaLinha]]:
    """Agrupa linhas por Matricula Chefia. Retorna dict[matricula_chefia] = lista."""
    grupos: Dict[str, List[DivergenciaLinha]] = defaultdict(list)

    col_map = {_normalizar_cabecalho(c): c for c in df.columns}

    def _col(nome_padrao: str) -> str:
        alvo = _normalizar_cabecalho(nome_padrao)
        return col_map.get(alvo, nome_padrao)

    for _, row in df.iterrows():
        mat_chefia = str(row.get(_col("Matricula Chefia"), "")).strip()
        if not mat_chefia:
            continue

        linha = DivergenciaLinha(
            nome_cr=str(row.get(_col("Nome CR"), "")).strip(),
            matricula=str(row.get(_col("Matricula"), "")).strip(),
            nome=str(row.get(_col("Nome"), "")).strip(),
            funcao=str(row.get(_col("Funcao"), "")).strip(),
            horario=str(row.get(_col("Descricao Horario"), "") or row.get(_col("Horario"), "")).strip(),
            data=str(row.get(_col("Data"), "")).strip(),
            dia=str(row.get(_col("DIA"), "")).strip(),
            batida=str(row.get(_col("Batida"), "")).strip(),
            situacao=str(row.get(_col("Situacao"), "")).strip(),
            descricao=str(row.get(_col("Descricao"), "")).strip(),
            quantidade=str(row.get(_col("Quantidade"), "")).strip(),
        )
        grupos[mat_chefia].append(linha)

    return dict(grupos)


def extrair_nomes_chefia(df: pd.DataFrame) -> Dict[str, str]:
    """Retorna dict[matricula_chefia] = nome_chefia (do próprio TXT)."""
    col_map = {_normalizar_cabecalho(c): c for c in df.columns}

    def _col(nome_padrao: str) -> str:
        alvo = _normalizar_cabecalho(nome_padrao)
        return col_map.get(alvo, nome_padrao)

    nomes: Dict[str, str] = {}
    for _, row in df.iterrows():
        mat = str(row.get(_col("Matricula Chefia"), "")).strip()
        nome = str(row.get(_col("Nome Chefia"), "")).strip()
        if mat and nome:
            nomes[mat] = nome
    return nomes


def extrair_periodo(df: pd.DataFrame) -> str:
    """Extrai período do arquivo (Data Inicial — Data Final)."""
    col_map = {_normalizar_cabecalho(c): c for c in df.columns}

    def _col(nome_padrao: str) -> str:
        alvo = _normalizar_cabecalho(nome_padrao)
        return col_map.get(alvo, nome_padrao)

    try:
        di = df[_col("Data Inicial")].iloc[0].strip()
        df_val = df[_col("Data Final")].iloc[0].strip()
        return f"{di} a {df_val}"
    except Exception:
        return "período não identificado"


# ─── Leitura da base de gestores ──────────────────────────────────────────────
def ler_base_gestores_texto(texto: str) -> Dict[str, Gestor]:
    """Lê base colada (Matricula Chefia;Nome;Email)."""
    linhas = [l.rstrip() for l in texto.splitlines() if l.strip()]
    if not linhas:
        raise ValueError("Nenhum dado foi colado na área de texto.")

    delim = "\t" if "\t" in linhas[0] else ";"
    cabecalhos = [c.strip() for c in linhas[0].split(delim)]

    # Mapear colunas
    cab_norm = {_normalizar_cabecalho(c): c for c in cabecalhos}

    aliases_matricula = ["MATRICULACHEFIA", "MATRICULA", "MATR", "REGISTRO", "CODIGO"]
    aliases_nome = ["NOMECHEFIA", "NOME", "NOMECOMPLETO", "GESTOR", "CHEFIA"]
    aliases_email = ["EMAIL", "EMAILALTERNATIVO", "EMAILGESTOR", "EMAILCHEFIA"]

    def _encontrar(possiveis):
        for p in possiveis:
            if p in cab_norm:
                return cab_norm[p]
        return None

    col_mat = _encontrar(aliases_matricula)
    col_nome = _encontrar(aliases_nome)
    col_email = _encontrar(aliases_email)

    if not col_mat:
        raise ValueError(
            f"Coluna de matrícula da chefia não encontrada. "
            f"Colunas detectadas: {', '.join(cabecalhos)}"
        )
    if not col_email:
        raise ValueError(
            f"Coluna de e-mail não encontrada. "
            f"Colunas detectadas: {', '.join(cabecalhos)}"
        )

    dados = []
    for linha in linhas[1:]:
        partes = [p.strip() for p in linha.split(delim)]
        if len(partes) < len(cabecalhos):
            partes += [""] * (len(cabecalhos) - len(partes))
        dados.append(partes[:len(cabecalhos)])

    df = pd.DataFrame(dados, columns=cabecalhos)
    base: Dict[str, Gestor] = {}

    for _, row in df.fillna("").iterrows():
        mat = str(row.get(col_mat, "")).strip()
        if not mat:
            continue
        base[mat] = Gestor(
            matricula_chefia=mat,
            nome=str(row.get(col_nome, "")).strip() if col_nome else "",
            email=str(row.get(col_email, "")).strip(),
        )

    if not base:
        raise ValueError("Nenhum gestor válido foi identificado na base colada.")
    return base


def ler_base_gestores_arquivo(caminho: Path) -> Dict[str, Gestor]:
    """Lê base de gestores de arquivo CSV/XLSX."""
    ext = caminho.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(caminho, dtype=str, sep=None, engine="python",
                         encoding="utf-8", keep_default_na=False)
    elif ext in [".xls", ".xlsx", ".xlsm"]:
        df = pd.read_excel(caminho, dtype=str).fillna("")
    else:
        raise ValueError("Formato não suportado. Use CSV ou XLSX.")

    # Converter para texto tabulado e reaproveitar
    linhas_texto = []
    linhas_texto.append("\t".join(str(c) for c in df.columns))
    for _, row in df.iterrows():
        linhas_texto.append("\t".join(str(row[c]) for c in df.columns))
    return ler_base_gestores_texto("\n".join(linhas_texto))


# ─── Montar corpo do e-mail ───────────────────────────────────────────────────
def montar_corpo_texto(
    nome_chefia: str,
    periodo: str,
    divergencias: List[DivergenciaLinha],
    header_template: str,
    footer_template: str,
) -> str:
    """Monta corpo do e-mail em texto simples com tabela formatada."""

    header = header_template.format(
        nome_chefia=nome_chefia,
        periodo=periodo,
    )

    # Agrupa por colaborador
    por_colab: Dict[str, List[DivergenciaLinha]] = defaultdict(list)
    for d in divergencias:
        chave = f"{d.matricula}|{d.nome}"
        por_colab[chave].append(d)

    corpo_tabela = ""
    sep = "-" * 110

    for chave, lista in sorted(por_colab.items(), key=lambda x: x[0].split("|")[1]):
        mat, nome_colab = chave.split("|", 1)
        funcao = lista[0].funcao if lista else ""
        corpo_tabela += f"\n{'='*110}\n"
        corpo_tabela += f"  Colaborador: {nome_colab}  |  Matrícula: {mat}  |  Função: {funcao}\n"
        corpo_tabela += f"{'='*110}\n"

        # Cabeçalho da tabela
        corpo_tabela += (
            f"  {'Data':<12} {'Dia':<5} {'Horário Previsto':<22} "
            f"{'Batidas':<28} {'Descrição':<28} {'Qtd':<6} {'Observação':<30}\n"
        )
        corpo_tabela += f"  {sep}\n"

        for d in sorted(lista, key=lambda x: x.data):
            corpo_tabela += (
                f"  {d.data:<12} {d.dia:<5} {d.horario:<22} "
                f"{d.batida:<28} {d.descricao:<28} {d.quantidade:<6} {'':<30}\n"
            )

        corpo_tabela += f"\n  Total de ocorrências: {len(lista)}\n"

    footer = footer_template.format(
        nome_chefia=nome_chefia,
        periodo=periodo,
    )

    return header + corpo_tabela + footer


def montar_corpo_html(
    nome_chefia: str,
    periodo: str,
    divergencias: List[DivergenciaLinha],
    header_template: str,
    footer_template: str,
) -> str:
    """Monta corpo do e-mail em HTML com tabela estilizada."""

    header_text = header_template.format(
        nome_chefia=nome_chefia,
        periodo=periodo,
    )
    header_html = header_text.replace("\n", "<br>")

    # Agrupa por colaborador
    por_colab: Dict[str, List[DivergenciaLinha]] = defaultdict(list)
    for d in divergencias:
        chave = f"{d.matricula}|{d.nome}"
        por_colab[chave].append(d)

    tabelas_html = ""
    style_table = (
        'style="border-collapse:collapse;width:100%;font-family:Calibri,Arial,sans-serif;'
        'font-size:11pt;margin-bottom:18px;"'
    )
    style_th = (
        'style="background-color:#4472C4;color:#fff;padding:6px 10px;'
        'border:1px solid #2F5496;text-align:left;font-size:10pt;"'
    )
    style_td = (
        'style="padding:5px 10px;border:1px solid #B4C6E7;font-size:10pt;"'
    )
    style_td_alt = (
        'style="padding:5px 10px;border:1px solid #B4C6E7;'
        'background-color:#D6E4F0;font-size:10pt;"'
    )
    style_colab = (
        'style="background-color:#2F5496;color:#fff;padding:8px 12px;'
        'font-size:11pt;font-weight:bold;margin-top:14px;"'
    )
    style_obs = (
        'style="padding:5px 10px;border:1px solid #B4C6E7;'
        'background-color:#FFF2CC;font-size:10pt;min-width:180px;"'
    )
    style_obs_alt = (
        'style="padding:5px 10px;border:1px solid #B4C6E7;'
        'background-color:#FFE699;font-size:10pt;min-width:180px;"'
    )

    for chave, lista in sorted(por_colab.items(), key=lambda x: x[0].split("|")[1]):
        mat, nome_colab = chave.split("|", 1)
        funcao = lista[0].funcao if lista else ""

        tabelas_html += f'<div {style_colab}>'
        tabelas_html += f'{nome_colab} &nbsp;|&nbsp; Mat: {mat} &nbsp;|&nbsp; {funcao}'
        tabelas_html += f'</div>\n'
        tabelas_html += f'<table {style_table}>\n'
        tabelas_html += '<tr>'
        for col in ["Data", "Dia", "Horário Previsto", "Batidas", "Descrição", "Qtd"]:
            tabelas_html += f'<th {style_th}>{col}</th>'
        style_th_obs = (
            'style="background-color:#BF8F00;color:#fff;padding:6px 10px;'
            'border:1px solid #806000;text-align:left;font-size:10pt;min-width:180px;"'
        )
        tabelas_html += f'<th {style_th_obs}>Observação do Gestor</th>'
        tabelas_html += '</tr>\n'

        for i, d in enumerate(sorted(lista, key=lambda x: x.data)):
            td = style_td_alt if i % 2 else style_td
            obs = style_obs_alt if i % 2 else style_obs
            tabelas_html += '<tr>'
            tabelas_html += f'<td {td}>{d.data}</td>'
            tabelas_html += f'<td {td}>{d.dia}</td>'
            tabelas_html += f'<td {td}>{d.horario}</td>'
            tabelas_html += f'<td {td}>{d.batida}</td>'
            tabelas_html += f'<td {td}>{d.descricao}</td>'
            tabelas_html += f'<td {td}>{d.quantidade}</td>'
            tabelas_html += f'<td {obs}>&nbsp;</td>'
            tabelas_html += '</tr>\n'

        tabelas_html += '</table>\n'
        tabelas_html += (
            f'<p style="font-size:10pt;color:#333;margin:2px 0 16px 4px;">'
            f'Total de ocorrências: <b>{len(lista)}</b></p>\n'
        )

    footer_text = footer_template.format(
        nome_chefia=nome_chefia,
        periodo=periodo,
    )
    footer_html = footer_text.replace("\n", "<br>")

    html = (
        '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">'
        f'{header_html}'
        f'{tabelas_html}'
        f'{footer_html}'
        '</div>'
    )
    return html


# ─── Integração Outlook ───────────────────────────────────────────────────────
def outlook_criar_rascunho(
    para: str,
    assunto: str,
    corpo_html: str,
) -> None:
    """Cria rascunho no Outlook com corpo HTML."""
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError(
            "win32com.client não disponível. Instale pywin32 para integração com Outlook."
        ) from e

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = para
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    mail.Save()


# ─── Log CSV ──────────────────────────────────────────────────────────────────
def escrever_log(log_path: Path, resultados: List[ResultadoProcesso]) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    existe = log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        if not existe:
            writer.writerow([
                "timestamp", "matricula_chefia", "nome_chefia", "email",
                "qtd_colaboradores", "qtd_divergencias", "status", "detalhe",
            ])
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        for r in resultados:
            writer.writerow([
                ts, r.matricula_chefia, r.nome_chefia, r.email,
                r.qtd_colaboradores, r.qtd_divergencias, r.status, r.detalhe,
            ])


# ─── Processamento principal ──────────────────────────────────────────────────
def processar_divergencias(
    txt_path: Path,
    base_gestores: Dict[str, Gestor],
    nomes_chefia_txt: Dict[str, str],
    criar_rascunho: bool = True,
    usar_html: bool = True,
    subject_template: str = SUBJECT_TEMPLATE,
    header_template: str = BODY_HEADER_TEMPLATE,
    footer_template: str = BODY_FOOTER_TEMPLATE,
    on_progress: Optional[Callable[[int, int], None]] = None,
    on_log: Optional[Callable[[str], None]] = None,
) -> Tuple[List[ResultadoProcesso], Path]:

    df = ler_txt_ponto(txt_path)
    grupos = agrupar_por_chefia(df)
    periodo = extrair_periodo(df)

    resultados: List[ResultadoProcesso] = []
    total = len(grupos)

    for i, (mat_chefia, divergencias) in enumerate(sorted(grupos.items()), start=1):
        if on_progress:
            on_progress(i, total)

        nome_chefia = nomes_chefia_txt.get(mat_chefia, mat_chefia)

        # Contar colaboradores únicos
        colabs_unicos = set(d.matricula for d in divergencias)

        if not criar_rascunho:
            # Modo visualização
            r = ResultadoProcesso(
                matricula_chefia=mat_chefia,
                nome_chefia=nome_chefia,
                email="",
                qtd_colaboradores=len(colabs_unicos),
                qtd_divergencias=len(divergencias),
                status="OK",
                detalhe=f"Visualização: {len(colabs_unicos)} colaborador(es), {len(divergencias)} divergência(s).",
            )
            resultados.append(r)
            if on_log:
                on_log(
                    f"OK: Chefia {mat_chefia} ({nome_chefia}) | "
                    f"{len(colabs_unicos)} colab(s), {len(divergencias)} diverg."
                )
            continue

        # Modo rascunho — procurar gestor na base
        gestor = base_gestores.get(mat_chefia)
        if not gestor:
            r = ResultadoProcesso(
                matricula_chefia=mat_chefia,
                nome_chefia=nome_chefia,
                email="",
                qtd_colaboradores=len(colabs_unicos),
                qtd_divergencias=len(divergencias),
                status="PULADO",
                detalhe="Matrícula da chefia não encontrada na base de gestores.",
            )
            resultados.append(r)
            if on_log:
                on_log(f"PULADO: Chefia {mat_chefia} ({nome_chefia}) — não está na base")
            continue

        if not gestor.email or not validar_email_basico(gestor.email):
            r = ResultadoProcesso(
                matricula_chefia=mat_chefia,
                nome_chefia=nome_chefia,
                email=gestor.email,
                qtd_colaboradores=len(colabs_unicos),
                qtd_divergencias=len(divergencias),
                status="FALHA",
                detalhe=f"E-mail inválido ou vazio: '{gestor.email}'",
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: Chefia {mat_chefia} ({nome_chefia}) — e-mail inválido")
            continue

        # Montar e-mail
        try:
            nome_final = gestor.nome or nome_chefia

            assunto = subject_template.format(
                periodo=periodo,
                nome_chefia=nome_final,
                matricula_chefia=mat_chefia,
            )

            if usar_html:
                corpo = montar_corpo_html(
                    nome_chefia=nome_final,
                    periodo=periodo,
                    divergencias=divergencias,
                    header_template=header_template,
                    footer_template=footer_template,
                )
            else:
                corpo = montar_corpo_texto(
                    nome_chefia=nome_final,
                    periodo=periodo,
                    divergencias=divergencias,
                    header_template=header_template,
                    footer_template=footer_template,
                )

            outlook_criar_rascunho(
                para=gestor.email,
                assunto=assunto,
                corpo_html=corpo if usar_html else corpo.replace("\n", "<br>"),
            )

            r = ResultadoProcesso(
                matricula_chefia=mat_chefia,
                nome_chefia=nome_final,
                email=gestor.email,
                qtd_colaboradores=len(colabs_unicos),
                qtd_divergencias=len(divergencias),
                status="OK",
                detalhe="Rascunho criado com sucesso.",
            )
            resultados.append(r)
            if on_log:
                on_log(
                    f"OK: Chefia {mat_chefia} ({nome_final}) | "
                    f"email={gestor.email} | {len(colabs_unicos)} colab(s) | rascunho salvo"
                )

        except Exception as e:
            r = ResultadoProcesso(
                matricula_chefia=mat_chefia,
                nome_chefia=nome_chefia,
                email=gestor.email,
                qtd_colaboradores=len(colabs_unicos),
                qtd_divergencias=len(divergencias),
                status="FALHA",
                detalhe=f"Erro ao criar rascunho: {e}",
            )
            resultados.append(r)
            if on_log:
                on_log(f"FALHA: Chefia {mat_chefia} | {e}")

    log_path = txt_path.parent / LOG_NAME
    escrever_log(log_path, resultados)
    return resultados, log_path


def resumir_resultados(resultados: List[ResultadoProcesso], modo: str) -> str:
    total = len(resultados)
    ok = sum(1 for r in resultados if r.status == "OK")
    pulado = sum(1 for r in resultados if r.status == "PULADO")
    falha = sum(1 for r in resultados if r.status == "FALHA")
    total_diverg = sum(r.qtd_divergencias for r in resultados)
    total_colabs = sum(r.qtd_colaboradores for r in resultados)

    return (
        f"{APP_NAME} v{VERSAO}\n"
        f"Modo: {modo}\n"
        f"Gestores processados: {total}\n"
        f"Colaboradores com divergência: {total_colabs}\n"
        f"Total de divergências: {total_diverg}\n"
        f"OK: {ok} | PULADOS: {pulado} | FALHAS: {falha}"
    )


# ─── Interface gráfica ────────────────────────────────────────────────────────
if CTK_AVAILABLE:

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()
            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")

            self.title(f"{APP_NAME} | v{VERSAO}")
            self.geometry("1300x900")
            self.minsize(1100, 800)

            self.txt_path: Optional[Path] = None
            self.base_path: Optional[Path] = None
            self._df_cache: Optional[pd.DataFrame] = None

            self._montar_ui()

        # ── UI ────────────────────────────────────────────────────────────
        def _montar_ui(self):
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(4, weight=1)

            # ── Título ──
            frame_top = ctk.CTkFrame(self)
            frame_top.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
            frame_top.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                frame_top,
                text=f"{APP_NAME} | v{VERSAO}",
                font=ctk.CTkFont(size=22, weight="bold"),
            ).grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

            ctk.CTkLabel(
                frame_top,
                text=(
                    "Lê arquivo TXT de divergência de ponto, agrupa por gestor/chefia "
                    "e cria rascunhos individuais no Outlook com a tabela de divergências no corpo."
                ),
                font=ctk.CTkFont(size=13),
                justify="left",
            ).grid(row=1, column=0, padx=14, pady=(0, 12), sticky="w")

            # ── Seleção do TXT ──
            frame_txt = ctk.CTkFrame(self)
            frame_txt.grid(row=1, column=0, padx=16, pady=8, sticky="ew")
            frame_txt.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(frame_txt, text="Arquivo TXT de ponto").grid(
                row=0, column=0, padx=12, pady=10, sticky="w"
            )
            self.ent_txt = ctk.CTkEntry(frame_txt)
            self.ent_txt.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_txt, "")

            ctk.CTkButton(
                frame_txt, text="Selecionar TXT", width=160,
                command=self._selecionar_txt,
            ).grid(row=0, column=2, padx=(6, 12), pady=10)

            # ── Base de gestores ──
            frame_base = ctk.CTkFrame(self)
            frame_base.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
            frame_base.grid_columnconfigure(1, weight=1)
            frame_base.grid_rowconfigure(2, weight=1)

            ctk.CTkLabel(frame_base, text="Base por arquivo").grid(
                row=0, column=0, padx=12, pady=10, sticky="w"
            )
            self.ent_base = ctk.CTkEntry(frame_base)
            self.ent_base.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_base, "")

            ctk.CTkButton(
                frame_base, text="Selecionar Base", width=160,
                command=self._selecionar_base,
            ).grid(row=0, column=2, padx=(6, 12), pady=10)

            ctk.CTkLabel(
                frame_base,
                text=(
                    "Ou cole a base abaixo. Cabeçalhos aceitos: "
                    "Matricula Chefia (ou Matricula), Nome (ou Nome Chefia), Email"
                ),
            ).grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 6), sticky="w")

            self.txt_base = ctk.CTkTextbox(frame_base, height=110)
            self.txt_base.grid(row=2, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")
            self.txt_base.insert(
                "1.0",
                "Matricula Chefia;Nome;Email\n"
                "001;NOME SOBRENOME;nome1@empresa.com\n"
                "002;NOME SOBRENOME;nome2@empresa.com\n"
            )

            # ── Templates ──
            frame_msg = ctk.CTkFrame(self)
            frame_msg.grid(row=3, column=0, padx=16, pady=8, sticky="ew")
            frame_msg.grid_columnconfigure(1, weight=1)
            frame_msg.grid_rowconfigure(3, weight=1)

            ctk.CTkLabel(frame_msg, text="Assunto").grid(
                row=0, column=0, padx=12, pady=6, sticky="w"
            )
            self.ent_assunto = ctk.CTkEntry(frame_msg)
            self.ent_assunto.grid(row=0, column=1, columnspan=2, padx=12, pady=6, sticky="ew")
            self.ent_assunto.insert(0, SUBJECT_TEMPLATE)

            ctk.CTkLabel(frame_msg, text="Cabeçalho").grid(
                row=1, column=0, padx=12, pady=(6, 4), sticky="nw"
            )
            self.txt_header = ctk.CTkTextbox(frame_msg, height=90)
            self.txt_header.grid(row=1, column=1, columnspan=2, padx=12, pady=(6, 4), sticky="ew")
            self.txt_header.insert("1.0", BODY_HEADER_TEMPLATE)

            ctk.CTkLabel(frame_msg, text="Rodapé").grid(
                row=2, column=0, padx=12, pady=(4, 6), sticky="nw"
            )
            self.txt_footer = ctk.CTkTextbox(frame_msg, height=80)
            self.txt_footer.grid(row=2, column=1, columnspan=2, padx=12, pady=(4, 6), sticky="ew")
            self.txt_footer.insert("1.0", BODY_FOOTER_TEMPLATE)

            self.chk_html_var = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                frame_msg, text="Enviar em HTML (tabela formatada)",
                variable=self.chk_html_var,
            ).grid(row=3, column=1, padx=12, pady=(0, 8), sticky="w")

            # ── Botões e Log ──
            frame_exec = ctk.CTkFrame(self)
            frame_exec.grid(row=4, column=0, padx=16, pady=8, sticky="nsew")
            frame_exec.grid_columnconfigure(0, weight=1)
            frame_exec.grid_rowconfigure(2, weight=1)

            botoes = ctk.CTkFrame(frame_exec)
            botoes.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

            self.btn_visualizar = ctk.CTkButton(
                botoes, text="Visualizar Split (sem enviar)",
                width=230, command=self._iniciar_visualizacao,
            )
            self.btn_visualizar.grid(row=0, column=0, padx=(0, 10), pady=8, sticky="w")

            self.btn_criar = ctk.CTkButton(
                botoes, text="Criar Rascunhos no Outlook",
                width=240, command=self._iniciar_rascunho,
            )
            self.btn_criar.grid(row=0, column=1, padx=10, pady=8, sticky="w")

            self.btn_validar = ctk.CTkButton(
                botoes, text="Validar Base Colada",
                width=180, command=self._validar_base,
            )
            self.btn_validar.grid(row=0, column=2, padx=10, pady=8, sticky="w")

            self.btn_reiniciar = ctk.CTkButton(
                botoes, text="Reiniciar",
                width=140, command=self._reiniciar,
            )
            self.btn_reiniciar.grid(row=0, column=3, padx=10, pady=8, sticky="w")

            self.btn_template = ctk.CTkButton(
                botoes, text="Gerar Template Excel",
                width=190, command=self._gerar_template_excel,
            )
            self.btn_template.grid(row=0, column=4, padx=10, pady=8, sticky="w")

            # Progresso
            self.progress = ctk.CTkProgressBar(frame_exec)
            self.progress.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="ew")
            self.progress.set(0)
            self.lbl_prog = ctk.CTkLabel(frame_exec, text="0/0")
            self.lbl_prog.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="e")

            # Log
            ctk.CTkLabel(
                frame_exec, text="Log do processamento",
                font=ctk.CTkFont(size=14, weight="bold"),
            ).grid(row=2, column=0, padx=12, pady=(4, 4), sticky="nw")

            self.txt_log = ctk.CTkTextbox(frame_exec)
            self.txt_log.grid(row=2, column=0, padx=12, pady=(30, 12), sticky="nsew")
            self.txt_log.configure(state="disabled")

            # Footer
            ctk.CTkLabel(
                self, text=FOOTER_UI, font=ctk.CTkFont(size=11),
            ).grid(row=5, column=0, padx=16, pady=(0, 10), sticky="e")

        # ── Helpers UI ────────────────────────────────────────────────────
        def _set_entry_readonly(self, entry, value: str):
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, value)
            entry.configure(state="readonly")

        def _log(self, msg: str):
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")

        def _limpar_log(self):
            self.txt_log.configure(state="normal")
            self.txt_log.delete("1.0", "end")
            self.txt_log.configure(state="disabled")

        def _on_progress(self, atual: int, total: int):
            if total <= 0:
                self.progress.set(0)
                self.lbl_prog.configure(text="0/0")
                return
            self.progress.set(atual / total)
            self.lbl_prog.configure(text=f"{atual}/{total}")
            self.update_idletasks()

        def _alternar_botoes(self, habilitar: bool):
            estado = "normal" if habilitar else "disabled"
            for btn in [self.btn_visualizar, self.btn_criar,
                        self.btn_validar, self.btn_reiniciar, self.btn_template]:
                btn.configure(state=estado)

        def _texto_base(self) -> str:
            texto = self.txt_base.get("1.0", "end").strip()
            exemplo_inicio = "Matricula Chefia;Nome;Email"
            if texto.startswith(exemplo_inicio) and "fabio@empresa.com" in texto:
                return ""
            return texto

        # ── Ações ─────────────────────────────────────────────────────────
        def _selecionar_txt(self):
            p = filedialog.askopenfilename(
                title="Selecione o TXT de divergência de ponto",
                filetypes=[("TXT", "*.txt"), ("CSV", "*.csv"), ("Todos", "*.*")],
            )
            if p:
                self.txt_path = Path(p)
                self._set_entry_readonly(self.ent_txt, str(self.txt_path))
                self._df_cache = None
                self._log(f"TXT selecionado: {self.txt_path}")

                # Pré-carregar para mostrar info
                try:
                    df = ler_txt_ponto(self.txt_path)
                    self._df_cache = df
                    grupos = agrupar_por_chefia(df)
                    nomes = extrair_nomes_chefia(df)
                    periodo = extrair_periodo(df)
                    total_linhas = len(df)
                    total_gestores = len(grupos)

                    self._log(f"Período: {periodo}")
                    self._log(f"Total de linhas: {total_linhas}")
                    self._log(f"Gestores encontrados: {total_gestores}")
                    self._log("")
                    for mat, divs in sorted(grupos.items()):
                        nome = nomes.get(mat, "?")
                        colabs = set(d.matricula for d in divs)
                        self._log(
                            f"  Chefia {mat} ({nome}): "
                            f"{len(colabs)} colaborador(es), {len(divs)} divergência(s)"
                        )
                    self._log("")
                except Exception as e:
                    self._log(f"Erro ao pré-carregar TXT: {e}")

        def _selecionar_base(self):
            p = filedialog.askopenfilename(
                title="Selecione a base de gestores",
                filetypes=[("Bases", "*.csv *.xls *.xlsx *.xlsm"), ("Todos", "*.*")],
            )
            if p:
                self.base_path = Path(p)
                self._set_entry_readonly(self.ent_base, str(self.base_path))
                self._log(f"Base selecionada: {self.base_path}")

        def _validar_base(self):
            try:
                texto = self._texto_base()
                if texto:
                    base = ler_base_gestores_texto(texto)
                elif self.base_path:
                    base = ler_base_gestores_arquivo(self.base_path)
                else:
                    raise ValueError("Cole a base ou selecione um arquivo.")

                self._log(f"Base validada: {len(base)} gestor(es)")
                msg = "Base validada com sucesso.\n\n"
                for mat, g in sorted(base.items()):
                    msg += f"  {mat} | {g.nome} | {g.email}\n"
                messagebox.showinfo("Base validada", msg)
            except Exception as e:
                self._log(f"Erro na base: {e}")
                messagebox.showerror("Erro", str(e))

        def _reiniciar(self):
            self.txt_path = None
            self.base_path = None
            self._df_cache = None
            self._set_entry_readonly(self.ent_txt, "")
            self._set_entry_readonly(self.ent_base, "")
            self._limpar_log()
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")
            self._log("Tela reiniciada.")

        def _gerar_template_excel(self):
            try:
                destino = Path.cwd() / "template_base_gestores.xlsx"
                dados = {
                    "Matricula Chefia": ["19", "1184", "207", "36"],
                    "Nome": [
                        "APARECIDO RONALDO PALAZON",
                        "FABIO EDUARDO NAKAMOTO",
                        "FLAVIO FRANCISCO DA SILVA",
                        "LUIZ CARLOS DA SILVA CAMARGO",
                    ],
                    "Email": [
                        "aparecido@empresa.com",
                        "fabio@empresa.com",
                        "flavio@empresa.com",
                        "luiz@empresa.com",
                    ],
                }
                df = pd.DataFrame(dados)
                df.to_excel(destino, sheet_name="Gestores", index=False, engine="openpyxl")
                self._log(f"Template gerado: {destino}")
                messagebox.showinfo("Template criado", f"Arquivo salvo em:\n{destino}")
            except Exception as e:
                messagebox.showerror("Erro", str(e))

        def _executar(self, criar_rascunho: bool):
            if not self.txt_path:
                messagebox.showerror("Erro", "Selecione o arquivo TXT de ponto.")
                return

            base_gestores: Dict[str, Gestor] = {}
            nomes_chefia_txt: Dict[str, str] = {}

            if criar_rascunho:
                texto = self._texto_base()
                if not texto and not self.base_path:
                    messagebox.showerror(
                        "Erro",
                        "Para criar rascunhos, cole a base de gestores ou selecione um arquivo.",
                    )
                    return
                try:
                    if texto:
                        base_gestores = ler_base_gestores_texto(texto)
                    else:
                        base_gestores = ler_base_gestores_arquivo(self.base_path)
                except Exception as e:
                    messagebox.showerror("Erro na base", str(e))
                    return

            # Extrair nomes do TXT
            try:
                df = self._df_cache or ler_txt_ponto(self.txt_path)
                nomes_chefia_txt = extrair_nomes_chefia(df)
            except Exception:
                pass

            self._alternar_botoes(False)
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")

            modo = "Criar Rascunhos" if criar_rascunho else "Visualizar Split"
            self._log("")
            self._log(f"Iniciando — Modo: {modo}")

            assunto_tpl = self.ent_assunto.get().strip() or SUBJECT_TEMPLATE
            header_tpl = self.txt_header.get("1.0", "end").strip() or BODY_HEADER_TEMPLATE
            footer_tpl = self.txt_footer.get("1.0", "end").strip() or BODY_FOOTER_TEMPLATE
            usar_html = self.chk_html_var.get()

            def worker():
                try:
                    resultados, log_path = processar_divergencias(
                        txt_path=self.txt_path,
                        base_gestores=base_gestores,
                        nomes_chefia_txt=nomes_chefia_txt,
                        criar_rascunho=criar_rascunho,
                        usar_html=usar_html,
                        subject_template=assunto_tpl,
                        header_template=header_tpl,
                        footer_template=footer_tpl,
                        on_progress=lambda a, t: self.after(0, self._on_progress, a, t),
                        on_log=lambda m: self.after(0, self._log, m),
                    )
                    resumo = resumir_resultados(resultados, modo)
                    self.after(0, self._log, "")
                    self.after(0, self._log, resumo)
                    self.after(0, self._log, f"Log: {log_path}")
                    self.after(0, messagebox.showinfo, "Concluído", resumo)
                except Exception as e:
                    msg = f"Erro: {e}\n\n{traceback.format_exc()}"
                    self.after(0, self._log, msg)
                    self.after(0, messagebox.showerror, "Erro", msg)
                finally:
                    self.after(0, self._alternar_botoes, True)

            threading.Thread(target=worker, daemon=True).start()

        def _iniciar_visualizacao(self):
            self._executar(criar_rascunho=False)

        def _iniciar_rascunho(self):
            self._executar(criar_rascunho=True)

else:
    class App:
        def __init__(self):
            raise RuntimeError("customtkinter não está instalado.")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

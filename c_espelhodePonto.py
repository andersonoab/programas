from __future__ import annotations

import csv
import json
import os
import re
import time
import traceback
import threading
import unicodedata
from dataclasses import dataclass
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


APP_NAME = "CustomerThinker | PDF Protect + Outlook"
VERSAO = "3.5-PONTO-PIS"
SUBPASTA_PROTEGIDOS = "protegidos"
LOG_NAME = "log_envio_documentos.csv"
FOOTER_UI = "Anderson Marinho | Igarapé Digital"
RH_CHAMADO_URL = ""


SUBJECT_TEMPLATE = "Informe de Rendimentos {ano_base} - Matrícula {matricula}"

BODY_TEMPLATE = (
    "Olá {nome},\n\n"
    "Segue em anexo o seu Informe de Rendimentos (ano-base {ano_base}).\n"
    "O arquivo está protegido.\n"
    "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
    "Caso identifique divergência de valores/dados ou tenha dificuldade de acesso, "
    "abra um chamado para o RH por este canal.\n"
    "{url_chamado}\n\n"
    "Atenciosamente,\n"
    "Recursos Humanos\n"
)


TEMPLATES_EMAIL = {
    "Informe de Rendimento": {
        "assunto": SUBJECT_TEMPLATE,
        "corpo": BODY_TEMPLATE,
    },
    "Férias": {
        "assunto": "Aviso e Recibo de Férias - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo o seu documento de férias.\n"
            "O arquivo está protegido.\n"
            "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
            "Em caso de dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
    "Rescisão": {
        "assunto": "Documentos Rescisórios - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo a documentação rescisória correspondente.\n"
            "O arquivo está protegido.\n"
            "Senha: seu CPF (somente números, sem ponto e sem traço).\n\n"
            "Em caso de dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
    "Documento Diverso": {
        "assunto": "Documento RH - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo o documento correspondente.\n"
            "Caso tenha dúvidas, entre em contato com o RH.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
    "Ponto": {
        "assunto": "Espelho de Ponto - Matrícula {matricula}",
        "corpo": (
            "Olá {nome},\n\n"
            "Segue em anexo o seu espelho de ponto para conferência.\n\n"
            "Pedimos a gentileza de validar as marcações, justificativas, eventuais ajustes, "
            "atestados e ocorrências registradas.\n\n"
            "Caso identifique alguma divergência, por favor retorne ao RH ou ao gestor responsável "
            "conforme o fluxo interno definido.\n\n"
            "Atenciosamente,\n"
            "Recursos Humanos\n"
        ),
    },
}


@dataclass
class Colaborador:
    matricula: str
    cpf: str
    pis: str
    nome: str
    email: str


@dataclass
class IdentificacaoPDF:
    cpf_encontrado: str
    nome_encontrado: str
    origem_cpf: str
    cpf_nome_arquivo: str
    divergencia_nome_arquivo: bool
    total_cpfs_validos: int
    texto_extraido: str


@dataclass
class IdentificacaoPonto:
    pis_encontrado: str
    origem_pis: str
    texto_extraido: str


@dataclass
class ResultadoProcesso:
    pdf_original: str
    cpf_pdf: str
    nome_pdf: str
    cpf_nome_arquivo: str
    origem_cpf: str
    encontrado_base: bool
    matricula: str
    email: str
    pdf_protegido: str
    status: str
    detalhe: str


def normalizar_texto_simples(valor: str) -> str:
    valor = str(valor or "").strip().upper()
    valor = unicodedata.normalize("NFKD", valor)
    valor = "".join(ch for ch in valor if not unicodedata.combining(ch))
    valor = re.sub(r"[^A-Z0-9 ]", " ", valor)
    valor = re.sub(r"\s+", " ", valor).strip()
    return valor


def normalizar_cpf(valor) -> str:
    if valor is None:
        return ""

    s = re.sub(r"\D", "", str(valor).strip())

    if not s:
        return ""

    if len(s) < 11:
        s = s.zfill(11)

    return s


def normalizar_pis(valor) -> str:
    if valor is None:
        return ""

    s = re.sub(r"\D", "", str(valor).strip())

    if not s:
        return ""

    if len(s) < 11:
        s = s.zfill(11)

    return s


def formatar_cpf(cpf: str) -> str:
    cpf = normalizar_cpf(cpf)

    if len(cpf) != 11:
        return cpf

    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"


def formatar_pis(pis: str) -> str:
    pis = normalizar_pis(pis)

    if len(pis) != 11:
        return pis

    return f"{pis[:3]}.{pis[3:8]}.{pis[8:10]}-{pis[10:]}"


def validar_cpf(cpf: str) -> bool:
    cpf = normalizar_cpf(cpf)

    if len(cpf) != 11:
        return False

    if cpf == cpf[0] * 11:
        return False

    soma1 = sum(int(cpf[i]) * (10 - i) for i in range(9))
    dig1 = (soma1 * 10) % 11
    dig1 = 0 if dig1 == 10 else dig1

    if dig1 != int(cpf[9]):
        return False

    soma2 = sum(int(cpf[i]) * (11 - i) for i in range(10))
    dig2 = (soma2 * 10) % 11
    dig2 = 0 if dig2 == 10 else dig2

    return dig2 == int(cpf[10])


def validar_pis(pis: str) -> bool:
    pis = normalizar_pis(pis)

    if len(pis) != 11:
        return False

    if pis == pis[0] * 11:
        return False

    pesos = [3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma = sum(int(pis[i]) * pesos[i] for i in range(10))
    resto = soma % 11
    digito = 11 - resto

    if digito in (10, 11):
        digito = 0

    return digito == int(pis[10])


def garantir_pasta(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def caminho_windows_estendido(p: Path) -> str:
    s = str(p)

    if os.name != "nt":
        return s

    s = str(p.resolve())

    if s.startswith("\\\\?\\"):
        return s

    if s.startswith("\\\\"):
        return "\\\\?\\UNC\\" + s.lstrip("\\")

    return "\\\\?\\" + s


def validar_email_basico(email: str) -> bool:
    if not email:
        return False

    e = email.strip()

    if " " in e or e.count("@") != 1:
        return False

    local, dom = e.split("@")

    if not local or not dom or "." not in dom:
        return False

    if not re.fullmatch(r"[A-Za-z0-9._%+\-]+", local):
        return False

    if not re.fullmatch(r"[A-Za-z0-9.\-]+", dom):
        return False

    tld = dom.rsplit(".", 1)[-1]

    if len(tld) < 2 or not re.fullmatch(r"[A-Za-z]{2,}", tld):
        return False

    return True


def extrair_ano_base_do_nome_arquivo(nome_arquivo: str) -> str:
    m = re.search(r"(20\d{2}|90\d{2})", nome_arquivo or "")

    if m:
        return m.group(1)

    ano_atual = int(time.strftime("%Y"))
    return str(ano_atual - 1)


def normalizar_nome_arquivo_manual(nome: str) -> str:
    nome = unicodedata.normalize("NFKD", str(nome or ""))
    nome = "".join(c for c in nome if not unicodedata.combining(c))
    nome = nome.upper()
    nome = re.sub(r"[^\w\s]", "", nome)
    nome = re.sub(r"\s+", "_", nome)
    nome = re.sub(r"_+", "_", nome)
    return nome.strip("_")


def limpar_nome_extraido(nome: str) -> str:
    nome = " ".join(str(nome or "").split()).strip(" -:")

    cortes = [
        " Natureza do Rendimento",
        " CPF",
        " CNPJ",
        " Valores",
        " Ano Calendário",
        " PIS",
        " PASEP",
        " NIT",
    ]

    for marcador in cortes:
        pos = nome.upper().find(marcador.upper())
        if pos > 0:
            nome = nome[:pos].strip(" -:")

    return nome


def limpar_nome_anexo_removendo_cpf(nome_arquivo: str, cpf11: str) -> str:
    base = nome_arquivo

    if cpf11:
        base = base.replace(cpf11, "")
        base = base.replace(formatar_cpf(cpf11), "")

    base = re.sub(r"__+", "_", base)
    base = re.sub(r"_\.", ".", base)
    base = re.sub(r"_-", "_", base)
    base = re.sub(r"_+", "_", base)
    base = re.sub(r"\s+", " ", base)
    base = base.strip("_").strip()

    return base or nome_arquivo


def limpar_nome_anexo_removendo_pis(nome_arquivo: str, pis11: str) -> str:
    base = nome_arquivo

    if pis11:
        base = base.replace(pis11, "")
        base = base.replace(formatar_pis(pis11), "")

    base = re.sub(r"__+", "_", base)
    base = re.sub(r"_\.", ".", base)
    base = re.sub(r"_+", "_", base)
    base = re.sub(r"\s+", " ", base)
    base = base.strip("_").strip()

    return base or nome_arquivo


def extrair_nome_do_arquivo_sem_cpf(nome_arquivo: str, cpf11: str = "") -> str:
    nome = Path(str(nome_arquivo or "")).stem

    if cpf11:
        nome = nome.replace(cpf11, "")
        nome = nome.replace(formatar_cpf(cpf11), "")

    nome = re.sub(r"(?i)_?protegido(?:_\d+)?$", "", nome)
    nome = re.sub(r"^[\s_\-\.]+", "", nome)
    nome = re.sub(r"[\s_\-\.]+$", "", nome)
    nome = re.sub(r"[_\-]+", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()

    return nome


def reduzir_nome_para_caminho(pasta: Path, nome_arquivo: str, margem_segura: int = 230) -> str:
    nome_arquivo = str(nome_arquivo or "").strip()

    if not nome_arquivo:
        return "arquivo.pdf"

    candidato = pasta / nome_arquivo

    if len(str(candidato)) <= margem_segura:
        return nome_arquivo

    stem = Path(nome_arquivo).stem
    suffix = Path(nome_arquivo).suffix or ".pdf"
    excesso = len(str(candidato)) - margem_segura
    novo_tamanho = max(20, len(stem) - excesso)
    stem = stem[:novo_tamanho].rstrip(" _.-")
    nome_reduzido = f"{stem}{suffix}"

    while len(str(pasta / nome_reduzido)) > margem_segura and len(stem) > 20:
        stem = stem[:-1].rstrip(" _.-")
        nome_reduzido = f"{stem}{suffix}"

    return nome_reduzido


def _extrair_texto_pypdf2(pdf_path: Path) -> str:
    try:
        try:
            from PyPDF2 import PdfReader
        except Exception:
            from pypdf import PdfReader
    except Exception:
        return ""

    textos = []

    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)

        for pagina in reader.pages:
            try:
                txt = pagina.extract_text() or ""
            except Exception:
                txt = ""

            textos.append(txt)

    return "\n".join(textos)


def _extrair_texto_pdfplumber(pdf_path: Path) -> str:
    try:
        import pdfplumber
    except Exception:
        return ""

    textos = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for pagina in pdf.pages:
                try:
                    txt = pagina.extract_text() or ""
                except Exception:
                    txt = ""

                textos.append(txt)
    except Exception:
        return ""

    return "\n".join(textos)


def extrair_texto_pdf(pdf_path: Path) -> str:
    texto1 = _extrair_texto_pypdf2(pdf_path)

    if texto1 and len(re.sub(r"\s+", "", texto1)) >= 20:
        return texto1

    texto2 = _extrair_texto_pdfplumber(pdf_path)

    if texto2 and len(re.sub(r"\s+", "", texto2)) > len(re.sub(r"\s+", "", texto1)):
        return texto2

    return texto1 or texto2 or ""


def remover_cnpjs_do_texto(texto: str) -> str:
    texto = texto or ""

    texto = re.sub(
        r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b",
        " ",
        texto
    )

    texto = re.sub(
        r"(?<!\d)\d{14}(?!\d)",
        " ",
        texto
    )

    texto = re.sub(
        r"\b\d{2}[\.\-\s/]*\d{3}[\.\-\s/]*\d{3}[\.\-\s/]*\d{4}[\.\-\s/]*\d{2}\b",
        " ",
        texto
    )

    return texto


def extrair_cpf_do_nome_arquivo(nome_arquivo: str) -> str:
    texto = remover_cnpjs_do_texto(nome_arquivo or "")

    candidatos = re.findall(
        r"\b\d{3}\.?\d{3}\.?\d{3}\-?\d{2}\b|(?<!\d)\d{11}(?!\d)",
        texto
    )

    for c in candidatos:
        cpf = normalizar_cpf(c)

        if validar_cpf(cpf):
            return cpf

    return ""


def extrair_pis_do_nome_arquivo(nome_arquivo: str) -> str:
    texto = remover_cnpjs_do_texto(nome_arquivo or "")

    candidatos = re.findall(
        r"\b\d{3}\.?\d{5}\.?\d{2}\-?\d{1}\b|(?<!\d)\d{11}(?!\d)",
        texto
    )

    for c in candidatos:
        pis = normalizar_pis(c)

        if validar_pis(pis):
            return pis

    return ""


def extrair_identidade_secao_beneficiario(texto: str) -> Tuple[str, str]:
    if not texto:
        return "", ""

    texto_sem_cnpj = remover_cnpjs_do_texto(texto)

    match_secao = re.search(
        r"2\.\s*PESSOA\s+F[IÍ]SICA\s+BENEFICI[ÁA]RIA\s+DOS\s+RENDIMENTOS",
        texto_sem_cnpj,
        flags=re.IGNORECASE,
    )

    chunks = []

    if match_secao:
        ini = max(0, match_secao.start() - 250)
        meio = match_secao.end()
        fim = min(len(texto_sem_cnpj), match_secao.end() + 1200)
        chunks.append(texto_sem_cnpj[meio:fim])
        chunks.append(texto_sem_cnpj[ini:fim])
    else:
        chunks.append(texto_sem_cnpj)

    padroes = [
        (r"CPF\s*:?\s*([\d\.\-]{11,14}).{0,180}?NOME(?:\s+COMPLETO)?\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,})", "cpf_nome"),
        (r"NOME(?:\s+COMPLETO)?\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,}).{0,180}?CPF\s*:?\s*([\d\.\-]{11,14})", "nome_cpf"),
        (r"([\d\.\-]{11,14})\s*CPF\s*:?\s*([A-ZÀ-Ú][A-ZÀ-Ú'\-\s]{5,})\s*NOME", "cpf_nome_fechando"),
    ]

    for chunk in chunks:
        for padrao, modo in padroes:
            m = re.search(padrao, chunk, flags=re.IGNORECASE | re.DOTALL)

            if not m:
                continue

            if modo == "cpf_nome":
                cpf, nome = m.group(1), m.group(2)
            elif modo == "nome_cpf":
                nome, cpf = m.group(1), m.group(2)
            else:
                cpf, nome = m.group(1), m.group(2)

            cpf = normalizar_cpf(cpf)
            nome = limpar_nome_extraido(nome)

            if validar_cpf(cpf):
                return cpf, nome

    return "", ""


def encontrar_cpfs_no_texto(texto: str) -> List[Dict[str, object]]:
    resultados: List[Dict[str, object]] = []

    if not texto:
        return resultados

    texto = remover_cnpjs_do_texto(texto)
    padrao = re.compile(r"\b\d{3}\.?\d{3}\.?\d{3}\-?\d{2}\b|(?<!\d)\d{11}(?!\d)")

    for match in padrao.finditer(texto):
        bruto = match.group(0)
        cpf = normalizar_cpf(bruto)

        if not validar_cpf(cpf):
            continue

        ini = max(0, match.start() - 160)
        fim = min(len(texto), match.end() + 160)
        contexto = texto[ini:fim]

        resultados.append(
            {
                "cpf": cpf,
                "cpf_bruto": bruto,
                "posicao": match.start(),
                "contexto": contexto,
            }
        )

    vistos = set()
    unicos: List[Dict[str, object]] = []

    for item in resultados:
        cpf = str(item["cpf"])

        if cpf not in vistos:
            vistos.add(cpf)
            unicos.append(item)

    return unicos


def encontrar_pis_no_texto(texto: str) -> List[Dict[str, object]]:
    resultados: List[Dict[str, object]] = []

    if not texto:
        return resultados

    texto = remover_cnpjs_do_texto(texto)

    padroes = [
        r"\bPIS\s*:?\s*([0-9][0-9\.\-\s/]{8,25}[0-9])",
        r"\bPASEP\s*:?\s*([0-9][0-9\.\-\s/]{8,25}[0-9])",
        r"\bNIT\s*:?\s*([0-9][0-9\.\-\s/]{8,25}[0-9])",
        r"\b\d{3}\.?\d{5}\.?\d{2}\-?\d{1}\b",
        r"(?<!\d)\d{11}(?!\d)",
        r"(?<!\d)\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d(?!\d)",
    ]

    for padrao in padroes:
        for match in re.finditer(padrao, texto, flags=re.IGNORECASE):
            if match.groups():
                bruto = match.group(1)
            else:
                bruto = match.group(0)

            pis = normalizar_pis(bruto)

            if len(pis) != 11:
                continue

            if not validar_pis(pis):
                continue

            ini = max(0, match.start() - 160)
            fim = min(len(texto), match.end() + 160)
            contexto = texto[ini:fim]

            resultados.append(
                {
                    "pis": pis,
                    "pis_bruto": bruto,
                    "posicao": match.start(),
                    "contexto": contexto,
                }
            )

    vistos = set()
    unicos: List[Dict[str, object]] = []

    for item in resultados:
        pis = str(item["pis"])

        if pis not in vistos:
            vistos.add(pis)
            unicos.append(item)

    return unicos


def pontuar_contexto_cpf(contexto: str) -> int:
    contexto_up = normalizar_texto_simples(contexto)
    score = 0

    termos_fortes = {
        "CPF": 20,
        "NOME": 10,
        "BENEFICIARIA": 12,
        "BENEFICIARIO": 12,
        "PESSOA FISICA": 12,
        "RENDIMENTOS": 8,
        "COLABORADOR": 8,
        "FUNCIONARIO": 8,
        "TITULAR": 8,
    }

    termos_negativos = {
        "RESPONSAVEL": -15,
        "CONTADOR": -15,
        "REPRESENTANTE": -12,
        "PROCURADOR": -10,
        "FONTE PAGADORA": -6,
        "CNPJ": -6,
    }

    for termo, peso in termos_fortes.items():
        if termo in contexto_up:
            score += peso

    for termo, peso in termos_negativos.items():
        if termo in contexto_up:
            score += peso

    if re.search(r"NOME\s*:?\s*.*?CPF", contexto, flags=re.IGNORECASE | re.DOTALL):
        score += 30

    if re.search(r"CPF\s*:?\s*\d", contexto, flags=re.IGNORECASE):
        score += 20

    if re.search(r"2\.?\s*PESSOA\s+F[IÍ]SICA\s+BENEFICI[ÁA]RIA", contexto, flags=re.IGNORECASE):
        score += 30

    return score


def pontuar_contexto_pis(contexto: str) -> int:
    contexto_up = normalizar_texto_simples(contexto)
    score = 0

    termos_fortes = {
        "PIS": 30,
        "PASEP": 30,
        "NIT": 25,
        "PONTO": 15,
        "ESPELHO": 12,
        "CARTAO": 12,
        "CARTAO DE PONTO": 20,
        "MATRICULA": 10,
        "COLABORADOR": 10,
        "FUNCIONARIO": 8,
        "NOME": 8,
    }

    termos_negativos = {
        "CNPJ": -20,
        "EMPRESA": -6,
        "FONTE PAGADORA": -8,
    }

    for termo, peso in termos_fortes.items():
        if termo in contexto_up:
            score += peso

    for termo, peso in termos_negativos.items():
        if termo in contexto_up:
            score += peso

    if re.search(r"(PIS|PASEP|NIT)\s*:?\s*\d", contexto, flags=re.IGNORECASE):
        score += 30

    return score


def escolher_cpf_mais_provavel(texto: str) -> Optional[Dict[str, object]]:
    candidatos = encontrar_cpfs_no_texto(texto)

    if not candidatos:
        return None

    for item in candidatos:
        item["score"] = pontuar_contexto_cpf(str(item["contexto"]))

    candidatos.sort(key=lambda x: (int(x["score"]), -int(x["posicao"])), reverse=True)
    return candidatos[0]


def escolher_pis_mais_provavel(texto: str) -> Optional[Dict[str, object]]:
    candidatos = encontrar_pis_no_texto(texto)

    if not candidatos:
        return None

    for item in candidatos:
        item["score"] = pontuar_contexto_pis(str(item["contexto"]))

    candidatos.sort(key=lambda x: (int(x["score"]), -int(x["posicao"])), reverse=True)
    return candidatos[0]


def extrair_nome_proximo_ao_cpf(texto: str, cpf: str) -> str:
    if not texto or not cpf:
        return ""

    cpf_fmt = formatar_cpf(cpf)

    padroes = [
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*{re.escape(cpf_fmt)}",
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*{re.escape(cpf)}",
        rf"Nome\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF\s*:?\s*\d{{3}}\.?\d{{3}}\.?\d{{3}}\-?\d{{2}}",
        rf"NOME COMPLETO\s*:?\s*([A-ZÀ-Úa-zà-ú'\-\s]+?)\s+CPF",
    ]

    for padrao in padroes:
        m = re.search(padrao, texto, flags=re.IGNORECASE | re.DOTALL)

        if m:
            nome = " ".join(m.group(1).split()).strip(" -:")

            if nome:
                return limpar_nome_extraido(nome)

    return ""


def identificar_pdf(pdf_path: Path) -> IdentificacaoPDF:
    texto = extrair_texto_pdf(pdf_path)
    cpf_nome_arquivo = extrair_cpf_do_nome_arquivo(pdf_path.name)

    cpf_secao, nome_secao = extrair_identidade_secao_beneficiario(texto)

    if cpf_secao:
        divergencia = bool(cpf_nome_arquivo and cpf_secao and cpf_nome_arquivo != cpf_secao)

        return IdentificacaoPDF(
            cpf_encontrado=cpf_secao,
            nome_encontrado=nome_secao,
            origem_cpf="secao_beneficiario",
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=divergencia,
            total_cpfs_validos=len(encontrar_cpfs_no_texto(texto)),
            texto_extraido=texto,
        )

    melhor = escolher_cpf_mais_provavel(texto)

    if melhor:
        cpf_texto = str(melhor["cpf"])
        nome_texto = extrair_nome_proximo_ao_cpf(texto, cpf_texto)
        origem = "conteudo_pdf"
        divergencia = bool(cpf_nome_arquivo and cpf_texto and cpf_nome_arquivo != cpf_texto)

        return IdentificacaoPDF(
            cpf_encontrado=cpf_texto,
            nome_encontrado=nome_texto,
            origem_cpf=origem,
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=divergencia,
            total_cpfs_validos=len(encontrar_cpfs_no_texto(texto)),
            texto_extraido=texto,
        )

    if cpf_nome_arquivo:
        return IdentificacaoPDF(
            cpf_encontrado=cpf_nome_arquivo,
            nome_encontrado="",
            origem_cpf="nome_arquivo_fallback",
            cpf_nome_arquivo=cpf_nome_arquivo,
            divergencia_nome_arquivo=False,
            total_cpfs_validos=0,
            texto_extraido=texto,
        )

    return IdentificacaoPDF(
        cpf_encontrado="",
        nome_encontrado="",
        origem_cpf="nao_encontrado",
        cpf_nome_arquivo=cpf_nome_arquivo,
        divergencia_nome_arquivo=False,
        total_cpfs_validos=0,
        texto_extraido=texto,
    )


def identificar_pdf_ponto_por_pis(pdf_path: Path) -> IdentificacaoPonto:
    texto = extrair_texto_pdf(pdf_path)
    pis_nome_arquivo = extrair_pis_do_nome_arquivo(pdf_path.name)

    if pis_nome_arquivo:
        return IdentificacaoPonto(
            pis_encontrado=pis_nome_arquivo,
            origem_pis="nome_arquivo",
            texto_extraido=texto,
        )

    melhor = escolher_pis_mais_provavel(texto)

    if melhor:
        return IdentificacaoPonto(
            pis_encontrado=str(melhor["pis"]),
            origem_pis="conteudo_pdf",
            texto_extraido=texto,
        )

    return IdentificacaoPonto(
        pis_encontrado="",
        origem_pis="nao_encontrado",
        texto_extraido=texto,
    )


def _normalizar_cabecalho(cab: str) -> str:
    return normalizar_texto_simples(cab).replace(" ", "")


def _mapear_colunas(df: pd.DataFrame, exigir_pis: bool = False) -> Dict[str, str]:
    aliases = {
        "cpf": ["CPF"],
        "pis": ["PIS", "PASEP", "NIT", "PIS/PASEP", "PIS PASEP", "Número PIS", "Numero PIS"],
        "matricula": ["Matrícula", "Matricula", "Matr", "Registro", "RE", "Código", "Codigo"],
        "nome": ["Nome", "Nome Completo", "Colaborador", "Funcionario", "Funcionário"],
        "email": ["Email Alternativo", "Email", "E-mail", "Email Pessoal", "E-mail Alternativo", "Correio Eletrônico"],
    }

    colunas_norm = {_normalizar_cabecalho(c): c for c in df.columns}
    resolvidas: Dict[str, str] = {}

    obrigatorias = ["matricula", "nome", "email"]

    if exigir_pis:
        obrigatorias.append("pis")
    else:
        obrigatorias.append("cpf")

    for chave in obrigatorias:
        encontrado = None

        for nome in aliases[chave]:
            alvo = _normalizar_cabecalho(nome)

            if alvo in colunas_norm:
                encontrado = colunas_norm[alvo]
                break

        if not encontrado:
            raise ValueError(
                "Não foi possível localizar as colunas obrigatórias na base. "
                f"Coluna ausente: {chave}. Colunas encontradas: {', '.join(map(str, df.columns))}"
            )

        resolvidas[chave] = encontrado

    for chave, possiveis in aliases.items():
        if chave in resolvidas:
            continue

        for nome in possiveis:
            alvo = _normalizar_cabecalho(nome)

            if alvo in colunas_norm:
                resolvidas[chave] = colunas_norm[alvo]
                break

    return resolvidas


def carregar_dataframe_base(arquivo_path: Path) -> pd.DataFrame:
    ext = arquivo_path.suffix.lower()

    if ext == ".csv":
        try:
            return pd.read_csv(
                arquivo_path,
                dtype=str,
                sep=None,
                engine="python",
                encoding="utf-8",
                keep_default_na=False,
            )
        except UnicodeDecodeError:
            return pd.read_csv(
                arquivo_path,
                dtype=str,
                sep=None,
                engine="python",
                encoding="latin1",
                keep_default_na=False,
            )

    if ext in [".xls", ".xlsx", ".xlsm"]:
        return pd.read_excel(arquivo_path, dtype=str).fillna("")

    if ext == ".xlsb":
        return pd.read_excel(arquivo_path, dtype=str, engine="pyxlsb").fillna("")

    raise ValueError("Formato não suportado. Use CSV/XLS/XLSX/XLSM/XLSB.")


def dataframe_base_texto(texto: str) -> pd.DataFrame:
    linhas = [linha.rstrip() for linha in texto.splitlines() if linha.strip()]

    if not linhas:
        raise ValueError("Nenhum dado foi colado na área de texto.")

    delimitador = "\t" if "\t" in linhas[0] else ";"
    cabecalhos = [c.strip() for c in linhas[0].split(delimitador)]
    dados = []

    for linha in linhas[1:]:
        partes = [p.strip() for p in linha.split(delimitador)]

        if len(partes) < len(cabecalhos):
            partes += [""] * (len(cabecalhos) - len(partes))

        dados.append(partes[: len(cabecalhos)])

    return pd.DataFrame(dados, columns=cabecalhos)


def _converter_dataframe_em_base(df: pd.DataFrame) -> Dict[str, Colaborador]:
    if df.empty:
        raise ValueError("A base está vazia.")

    mapa = _mapear_colunas(df, exigir_pis=False)
    base: Dict[str, Colaborador] = {}

    for _, row in df.fillna("").iterrows():
        cpf = normalizar_cpf(row.get(mapa["cpf"], ""))

        if not cpf:
            continue

        if not validar_cpf(cpf):
            continue

        pis = normalizar_pis(row.get(mapa.get("pis", ""), "")) if "pis" in mapa else ""

        base[cpf] = Colaborador(
            matricula=str(row.get(mapa["matricula"], "")).strip(),
            cpf=cpf,
            pis=pis,
            nome=str(row.get(mapa["nome"], "")).strip(),
            email=str(row.get(mapa["email"], "")).strip(),
        )

    if not base:
        raise ValueError("Nenhum CPF válido foi identificado na base.")

    return base


def _converter_dataframe_em_base_pis(df: pd.DataFrame) -> Dict[str, Colaborador]:
    if df.empty:
        raise ValueError("A base está vazia.")

    mapa = _mapear_colunas(df, exigir_pis=True)
    base: Dict[str, Colaborador] = {}

    for _, row in df.fillna("").iterrows():
        pis = normalizar_pis(row.get(mapa["pis"], ""))

        if not pis:
            continue

        if not validar_pis(pis):
            continue

        cpf = normalizar_cpf(row.get(mapa.get("cpf", ""), "")) if "cpf" in mapa else ""

        base[pis] = Colaborador(
            matricula=str(row.get(mapa["matricula"], "")).strip(),
            cpf=cpf,
            pis=pis,
            nome=str(row.get(mapa["nome"], "")).strip(),
            email=str(row.get(mapa["email"], "")).strip(),
        )

    if not base:
        raise ValueError("Nenhum PIS válido foi identificado na base.")

    return base


def ler_base_colaboradores_arquivo(arquivo_path: Path) -> Dict[str, Colaborador]:
    df = carregar_dataframe_base(arquivo_path)
    return _converter_dataframe_em_base(df)


def ler_base_colaboradores_texto(texto: str) -> Dict[str, Colaborador]:
    df = dataframe_base_texto(texto)
    return _converter_dataframe_em_base(df)


def ler_base_colaboradores_pis_arquivo(arquivo_path: Path) -> Dict[str, Colaborador]:
    df = carregar_dataframe_base(arquivo_path)
    return _converter_dataframe_em_base_pis(df)


def ler_base_colaboradores_pis_texto(texto: str) -> Dict[str, Colaborador]:
    df = dataframe_base_texto(texto)
    return _converter_dataframe_em_base_pis(df)


def obter_base(base_path: Optional[Path], base_texto: str) -> Dict[str, Colaborador]:
    if base_texto and base_texto.strip():
        return ler_base_colaboradores_texto(base_texto)

    if base_path:
        return ler_base_colaboradores_arquivo(base_path)

    raise ValueError("Informe uma base por arquivo ou cole os dados diretamente na tela.")


def obter_base_pis(base_path: Optional[Path], base_texto: str) -> Dict[str, Colaborador]:
    if base_texto and base_texto.strip():
        return ler_base_colaboradores_pis_texto(base_texto)

    if base_path:
        return ler_base_colaboradores_pis_arquivo(base_path)

    raise ValueError("Informe uma base por arquivo ou cole os dados diretamente na tela.")


def proteger_pdf_com_senha(pdf_path: Path, senha: str, saida_dir: Path) -> Path:
    try:
        import pikepdf
    except Exception as e:
        raise RuntimeError("pikepdf não está instalado. Instale com: pip install pikepdf") from e

    garantir_pasta(saida_dir)

    nome_base = extrair_nome_do_arquivo_sem_cpf(pdf_path.name, extrair_cpf_do_nome_arquivo(pdf_path.name)) or pdf_path.stem
    nome_base = normalizar_nome_arquivo_manual(nome_base) or normalizar_nome_arquivo_manual(pdf_path.stem) or "DOCUMENTO"

    out_name = f"{senha}_{nome_base}_protegido.pdf"
    out_name = reduzir_nome_para_caminho(saida_dir, out_name)
    out_path = saida_dir / out_name

    contador = 1

    while out_path.exists():
        out_name = reduzir_nome_para_caminho(saida_dir, f"{senha}_{nome_base}_protegido_{contador}.pdf")
        out_path = saida_dir / out_name
        contador += 1

    pdf_in_path = caminho_windows_estendido(pdf_path)
    pdf_out_path = caminho_windows_estendido(out_path)

    with pikepdf.open(pdf_in_path) as pdf:
        pdf.save(pdf_out_path, encryption=pikepdf.Encryption(owner=senha, user=senha, R=4))

    return out_path


def outlook_criar_rascunho_sem_exibir(
    para: str,
    assunto: str,
    corpo: str,
    anexo_path: Path,
    display_name: str,
    cc: str = "",
) -> None:
    try:
        import win32com.client as win32
    except Exception as e:
        raise RuntimeError(
            "win32com.client não está disponível. Instale pywin32 para integração com Outlook."
        ) from e

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = para
    if cc:
        mail.CC = cc
    mail.Subject = assunto
    mail.Body = corpo
    mail.Attachments.Add(str(anexo_path.resolve()), 1, 1, display_name)
    mail.Save()


def escrever_log_csv(log_path: Path, resultados: List[ResultadoProcesso]) -> None:
    garantir_pasta(log_path.parent)
    existe = log_path.exists()

    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")

        if not existe:
            writer.writerow(
                [
                    "timestamp",
                    "pdf_original",
                    "identificador_pdf",
                    "nome_pdf",
                    "identificador_nome_arquivo",
                    "origem_identificador",
                    "encontrado_base",
                    "matricula",
                    "email",
                    "pdf_anexado_ou_protegido",
                    "status",
                    "detalhe",
                ]
            )

        ts = time.strftime("%Y-%m-%d %H:%M:%S")

        for r in resultados:
            writer.writerow(
                [
                    ts,
                    r.pdf_original,
                    r.cpf_pdf,
                    r.nome_pdf,
                    r.cpf_nome_arquivo,
                    r.origem_cpf,
                    "SIM" if r.encontrado_base else "NAO",
                    r.matricula,
                    r.email,
                    r.pdf_protegido,
                    r.status,
                    r.detalhe,
                ]
            )


def resumir_resultados(resultados: List[ResultadoProcesso], modo: str) -> str:
    total = len(resultados)
    ok = sum(1 for r in resultados if r.status == "OK")
    pulado = sum(1 for r in resultados if r.status == "PULADO")
    falha = sum(1 for r in resultados if r.status == "FALHA")

    linhas = [
        f"{APP_NAME} v{VERSAO}",
        f"Modo: {modo}",
        f"Total PDFs: {total}",
        f"OK: {ok}",
        f"PULADOS: {pulado}",
        f"FALHAS: {falha}",
    ]

    return "\n".join(linhas)


def listar_pdfs(origem: Path) -> List[Path]:
    if origem.is_file():
        if origem.suffix.lower() != ".pdf":
            raise ValueError("O arquivo selecionado não é um PDF.")

        return [origem]

    if origem.is_dir():
        pdfs = sorted([p for p in origem.glob("*.pdf") if p.is_file()])

        if not pdfs:
            raise ValueError("Nenhum PDF encontrado na pasta selecionada.")

        return pdfs

    raise ValueError("Origem inválida para PDFs.")


def localizar_colaborador(base: Dict[str, Colaborador], cpf: str) -> Optional[Colaborador]:
    cpf = normalizar_cpf(cpf)

    if not cpf:
        return None

    return base.get(cpf)


def localizar_colaborador_por_pis(base: Dict[str, Colaborador], pis: str) -> Optional[Colaborador]:
    pis = normalizar_pis(pis)

    if not pis:
        return None

    return base.get(pis)


def processar_arquivos(
    origem_pdf: Path,
    base_path: Optional[Path] = None,
    base_texto: str = "",
    criar_rascunho: bool = True,
    on_progress: Optional[Callable[[int, int], None]] = None,
    on_log: Optional[Callable[[str], None]] = None,
    subject_template: Optional[str] = None,
    body_template: Optional[str] = None,
    cc_email: str = "",
) -> Tuple[List[ResultadoProcesso], Path]:

    pdfs = listar_pdfs(origem_pdf)
    base: Dict[str, Colaborador] = {}

    if criar_rascunho:
        base = obter_base(base_path, base_texto)

    pasta_saida = origem_pdf.parent if origem_pdf.is_file() else origem_pdf
    saida_dir = pasta_saida / SUBPASTA_PROTEGIDOS
    garantir_pasta(saida_dir)

    assunto_template_final = subject_template or SUBJECT_TEMPLATE
    corpo_template_final = body_template or BODY_TEMPLATE

    resultados: List[ResultadoProcesso] = []
    total = len(pdfs)

    for i, pdf in enumerate(pdfs, start=1):
        if on_progress:
            on_progress(i, total)

        try:
            identificacao = identificar_pdf(pdf)
        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf="erro_extracao",
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao ler PDF: {e}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

            continue

        cpf_pdf = normalizar_cpf(identificacao.cpf_encontrado)
        nome_pdf = identificacao.nome_encontrado
        detalhe_base = []

        if identificacao.divergencia_nome_arquivo:
            detalhe_base.append(
                f"CPF do conteúdo diverge do nome do arquivo ({formatar_cpf(identificacao.cpf_nome_arquivo)})."
            )

        if not cpf_pdf:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe="CPF não encontrado no conteúdo do PDF nem no nome do arquivo.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

            continue

        if not criar_rascunho:
            detalhe = f"CPF identificado por {identificacao.origem_cpf}."

            if nome_pdf:
                detalhe += f" Nome lido: {nome_pdf}."

            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="OK",
                detalhe=detalhe,
            )
            resultados.append(r)

            if on_log:
                on_log(f"OK: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} | origem={identificacao.origem_cpf}")

            continue

        colab = localizar_colaborador(base, cpf_pdf)

        if not colab:
            detalhe = "CPF encontrado no PDF, mas não está na base informada."

            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="PULADO",
                detalhe=detalhe,
            )
            resultados.append(r)

            if on_log:
                on_log(f"PULADO: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} não localizado na base")

            continue

        if not colab.email:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe="Email vazio na base para o CPF localizado.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail vazio na base")

            continue

        if not validar_email_basico(colab.email):
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=f"E-mail inválido na base: {colab.email}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail inválido | {colab.email}")

            continue

        try:
            ano_base = extrair_ano_base_do_nome_arquivo(pdf.name)

            assunto = assunto_template_final.format(
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                nome=(colab.nome or nome_pdf or "Colaborador(a)"),
                url_chamado=RH_CHAMADO_URL,
            )

            corpo = corpo_template_final.format(
                nome=(colab.nome or nome_pdf or "Colaborador(a)"),
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                url_chamado=RH_CHAMADO_URL,
            )

            protegido = proteger_pdf_com_senha(pdf, cpf_pdf, saida_dir)
            display_name = limpar_nome_anexo_removendo_cpf(pdf.name, cpf_pdf)

            outlook_criar_rascunho_sem_exibir(
                para=colab.email,
                assunto=assunto,
                corpo=corpo,
                anexo_path=protegido,
                display_name=display_name,
                cc=cc_email,
            )

            detalhe = "Rascunho criado com PDF protegido."

            if detalhe_base:
                detalhe += " " + " ".join(detalhe_base)

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido=protegido.name,
                status="OK",
                detalhe=detalhe,
            )
            resultados.append(r)

            if on_log:
                on_log(f"OK: {pdf.name} | CPF={formatar_cpf(cpf_pdf)} | email={colab.email} | rascunho salvo")

        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=cpf_pdf,
                nome_pdf=nome_pdf,
                cpf_nome_arquivo=identificacao.cpf_nome_arquivo,
                origem_cpf=identificacao.origem_cpf,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao proteger PDF ou criar rascunho: {e}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

    log_path = pasta_saida / LOG_NAME
    escrever_log_csv(log_path, resultados)

    return resultados, log_path


def processar_arquivos_ponto_por_pis(
    origem_pdf: Path,
    base_path: Optional[Path] = None,
    base_texto: str = "",
    criar_rascunho: bool = True,
    on_progress: Optional[Callable[[int, int], None]] = None,
    on_log: Optional[Callable[[str], None]] = None,
    subject_template: Optional[str] = None,
    body_template: Optional[str] = None,
    cc_email: str = "",
) -> Tuple[List[ResultadoProcesso], Path]:

    pdfs = listar_pdfs(origem_pdf)
    base: Dict[str, Colaborador] = {}

    if criar_rascunho:
        base = obter_base_pis(base_path, base_texto)

    pasta_saida = origem_pdf.parent if origem_pdf.is_file() else origem_pdf

    assunto_template_final = subject_template or TEMPLATES_EMAIL["Ponto"]["assunto"]
    corpo_template_final = body_template or TEMPLATES_EMAIL["Ponto"]["corpo"]

    resultados: List[ResultadoProcesso] = []
    total = len(pdfs)

    for i, pdf in enumerate(pdfs, start=1):
        if on_progress:
            on_progress(i, total)

        try:
            identificacao = identificar_pdf_ponto_por_pis(pdf)
            pis_pdf = normalizar_pis(identificacao.pis_encontrado)
            origem_pis = identificacao.origem_pis
        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf="erro_extracao_pis",
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao ler PDF para localizar PIS: {e}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

            continue

        if not pis_pdf:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf="",
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf="pis_nao_encontrado",
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe="PIS não encontrado no conteúdo do PDF nem no nome do arquivo.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | PIS não encontrado")

            continue

        if not criar_rascunho:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="OK",
                detalhe=f"PIS identificado por {origem_pis}: {formatar_pis(pis_pdf)}.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"OK: {pdf.name} | PIS={formatar_pis(pis_pdf)} | origem={origem_pis}")

            continue

        colab = localizar_colaborador_por_pis(base, pis_pdf)

        if not colab:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf="",
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=False,
                matricula="",
                email="",
                pdf_protegido="",
                status="PULADO",
                detalhe="PIS encontrado no PDF, mas não localizado na base informada.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"PULADO: {pdf.name} | PIS={formatar_pis(pis_pdf)} não localizado na base")

            continue

        if not colab.email:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf=colab.nome,
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=True,
                matricula=colab.matricula,
                email="",
                pdf_protegido="",
                status="FALHA",
                detalhe="Email vazio na base para o PIS localizado.",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail vazio na base")

            continue

        if not validar_email_basico(colab.email):
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf=colab.nome,
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=f"E-mail inválido na base: {colab.email}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | e-mail inválido | {colab.email}")

            continue

        try:
            ano_base = extrair_ano_base_do_nome_arquivo(pdf.name)

            assunto = assunto_template_final.format(
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                nome=(colab.nome or "Colaborador(a)"),
                url_chamado=RH_CHAMADO_URL,
                pis=formatar_pis(pis_pdf),
            )

            corpo = corpo_template_final.format(
                nome=(colab.nome or "Colaborador(a)"),
                ano_base=ano_base,
                matricula=(colab.matricula or "N/A"),
                url_chamado=RH_CHAMADO_URL,
                pis=formatar_pis(pis_pdf),
            )

            display_name = limpar_nome_anexo_removendo_pis(pdf.name, pis_pdf)

            outlook_criar_rascunho_sem_exibir(
                para=colab.email,
                assunto=assunto,
                corpo=corpo,
                anexo_path=pdf,
                display_name=display_name,
                cc=cc_email,
            )

            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf=colab.nome,
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido=pdf.name,
                status="OK",
                detalhe="Rascunho de ponto criado sem proteção por senha. PDF original anexado.",
            )
            resultados.append(r)

            if on_log:
                on_log(
                    f"OK: {pdf.name} | PIS={formatar_pis(pis_pdf)} | email={colab.email} | rascunho salvo sem senha"
                )

        except Exception as e:
            r = ResultadoProcesso(
                pdf_original=pdf.name,
                cpf_pdf=pis_pdf,
                nome_pdf=colab.nome,
                cpf_nome_arquivo="",
                origem_cpf=origem_pis,
                encontrado_base=True,
                matricula=colab.matricula,
                email=colab.email,
                pdf_protegido="",
                status="FALHA",
                detalhe=f"Erro ao criar rascunho de ponto sem senha: {e}",
            )
            resultados.append(r)

            if on_log:
                on_log(f"FALHA: {pdf.name} | {r.detalhe}")

    log_path = pasta_saida / LOG_NAME
    escrever_log_csv(log_path, resultados)

    return resultados, log_path


def gerar_template_excel() -> Path:
    caminho = Path.cwd() / "template_base_envio_documentos.xlsx"

    dados = [
        {
            "CPF": "12345678909",
            "PIS": "12345678901",
            "Matrícula": "123456",
            "Nome": "Nome e Sobrenome",
            "Email Alternativo": "email@exemplo.com",
        }
    ]

    df = pd.DataFrame(dados)
    df.to_excel(caminho, index=False)

    return caminho


if CTK_AVAILABLE:

    class App(ctk.CTk):
        def __init__(self):
            super().__init__()

            ctk.set_appearance_mode("System")
            ctk.set_default_color_theme("blue")

            self.title(f"{APP_NAME} | v{VERSAO}")
            self.geometry("1380x980")
            self.minsize(1280, 900)

            self.origem_pdf: Optional[Path] = None
            self.base_path: Optional[Path] = None
            self.templates_email = json.loads(json.dumps(TEMPLATES_EMAIL))
            self.tipo_documento_atual = "Informe de Rendimento"

            self._montar_ui()
            self._carregar_template_na_tela()

        def _montar_ui(self):
            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(4, weight=1)

            frame_top = ctk.CTkFrame(self)
            frame_top.grid(row=0, column=0, padx=16, pady=(14, 8), sticky="ew")
            frame_top.grid_columnconfigure(0, weight=1)

            lbl_titulo = ctk.CTkLabel(
                frame_top,
                text=f"{APP_NAME} | v{VERSAO}",
                font=ctk.CTkFont(size=22, weight="bold"),
            )
            lbl_titulo.grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

            lbl_sub = ctk.CTkLabel(
                frame_top,
                text=(
                    "Envio de documentos por Outlook. Para Informe/Férias/Rescisão, localiza CPF e protege com senha. "
                    "Para Ponto, localiza PIS e cria rascunho sem senha, anexando o PDF original."
                ),
                font=ctk.CTkFont(size=13),
                justify="left",
            )
            lbl_sub.grid(row=1, column=0, padx=14, pady=(0, 12), sticky="w")

            frame_origem = ctk.CTkFrame(self)
            frame_origem.grid(row=1, column=0, padx=16, pady=8, sticky="ew")
            frame_origem.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(frame_origem, text="Origem PDF").grid(row=0, column=0, padx=12, pady=10, sticky="w")

            self.ent_origem = ctk.CTkEntry(frame_origem)
            self.ent_origem.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_origem, "")

            ctk.CTkButton(frame_origem, text="Selecionar Pasta", width=140, command=self._selecionar_pasta).grid(
                row=0, column=2, padx=6, pady=10
            )

            ctk.CTkButton(frame_origem, text="Selecionar PDF", width=140, command=self._selecionar_pdf).grid(
                row=0, column=3, padx=(6, 12), pady=10
            )

            frame_base = ctk.CTkFrame(self)
            frame_base.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
            frame_base.grid_columnconfigure(1, weight=1)
            frame_base.grid_rowconfigure(2, weight=1)

            ctk.CTkLabel(frame_base, text="Base por arquivo").grid(row=0, column=0, padx=12, pady=10, sticky="w")

            self.ent_base = ctk.CTkEntry(frame_base)
            self.ent_base.grid(row=0, column=1, padx=12, pady=10, sticky="ew")
            self._set_entry_readonly(self.ent_base, "")

            ctk.CTkButton(frame_base, text="Selecionar Base", width=140, command=self._selecionar_base).grid(
                row=0, column=2, padx=(6, 12), pady=10
            )

            ctk.CTkLabel(
                frame_base,
                text=(
                    "Ou cole a base abaixo. Para Ponto, a coluna PIS/PASEP/NIT é obrigatória. "
                    "Cabeçalhos aceitos: CPF, PIS, PASEP, NIT, Matrícula/Matricula, Nome, Email/Email Alternativo."
                ),
            ).grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 6), sticky="w")

            self.txt_base = ctk.CTkTextbox(frame_base, height=130)
            self.txt_base.grid(row=2, column=0, columnspan=3, padx=12, pady=(0, 12), sticky="ew")
            self.txt_base.insert(
                "1.0",
                "CPF;PIS;Matrícula;Nome;Email Alternativo\n"
                "12345678909;12345678901;123456;Nome e Sobrenome;email@exemplo.com\n"
            )

            frame_msg = ctk.CTkFrame(self)
            frame_msg.grid(row=3, column=0, padx=16, pady=8, sticky="ew")
            frame_msg.grid_columnconfigure(1, weight=1)
            frame_msg.grid_rowconfigure(2, weight=1)

            ctk.CTkLabel(frame_msg, text="Tipo do documento").grid(row=0, column=0, padx=12, pady=10, sticky="w")

            self.cmb_tipo_documento = ctk.CTkComboBox(
                frame_msg,
                values=list(self.templates_email.keys()),
                command=self._ao_mudar_tipo_documento,
                width=260,
            )
            self.cmb_tipo_documento.grid(row=0, column=1, padx=12, pady=10, sticky="w")
            self.cmb_tipo_documento.set(self.tipo_documento_atual)

            ctk.CTkButton(frame_msg, text="Restaurar Modelo", width=150, command=self._restaurar_modelo_padrao).grid(
                row=0, column=2, padx=6, pady=10
            )

            ctk.CTkButton(frame_msg, text="Salvar Modelo Atual", width=170, command=self._salvar_modelo_atual).grid(
                row=0, column=3, padx=(6, 12), pady=10
            )

            ctk.CTkLabel(frame_msg, text="Assunto").grid(row=1, column=0, padx=12, pady=6, sticky="w")

            self.ent_assunto = ctk.CTkEntry(frame_msg)
            self.ent_assunto.grid(row=1, column=1, columnspan=3, padx=12, pady=6, sticky="ew")

            ctk.CTkLabel(frame_msg, text="CC / Cópia").grid(row=2, column=0, padx=12, pady=6, sticky="w")

            self.ent_cc = ctk.CTkEntry(frame_msg)
            self.ent_cc.grid(row=2, column=1, columnspan=3, padx=12, pady=6, sticky="ew")
            self.ent_cc.insert(0, "")

            ctk.CTkLabel(frame_msg, text="Corpo do e-mail").grid(row=3, column=0, padx=12, pady=(6, 12), sticky="nw")

            self.txt_corpo = ctk.CTkTextbox(frame_msg, height=170)
            self.txt_corpo.grid(row=3, column=1, columnspan=3, padx=12, pady=(6, 12), sticky="ew")

            frame_exec = ctk.CTkFrame(self)
            frame_exec.grid(row=4, column=0, padx=16, pady=8, sticky="nsew")
            frame_exec.grid_columnconfigure(0, weight=1)
            frame_exec.grid_rowconfigure(2, weight=1)

            botoes = ctk.CTkFrame(frame_exec)
            botoes.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")
            botoes.grid_columnconfigure(7, weight=1)

            self.btn_localizar = ctk.CTkButton(
                botoes,
                text="Encontrar CPF/PIS",
                width=180,
                command=self._iniciar_localizacao,
            )
            self.btn_localizar.grid(row=0, column=0, padx=(0, 10), pady=8, sticky="w")

            self.btn_iniciar = ctk.CTkButton(
                botoes,
                text="Criar Rascunhos",
                width=190,
                command=self._iniciar_rascunho,
            )
            self.btn_iniciar.grid(row=0, column=1, padx=10, pady=8, sticky="w")

            self.btn_validar_base = ctk.CTkButton(
                botoes,
                text="Validar Base",
                width=150,
                command=self._validar_base_colada,
            )
            self.btn_validar_base.grid(row=0, column=2, padx=10, pady=8, sticky="w")

            self.btn_reiniciar = ctk.CTkButton(
                botoes,
                text="Reiniciar",
                width=140,
                command=self._reiniciar,
            )
            self.btn_reiniciar.grid(row=0, column=3, padx=10, pady=8, sticky="w")

            self.btn_template = ctk.CTkButton(
                botoes,
                text="Gerar Template Excel",
                width=190,
                command=self._gerar_template_excel,
            )
            self.btn_template.grid(row=0, column=4, padx=10, pady=8, sticky="w")

            self.progress = ctk.CTkProgressBar(frame_exec)
            self.progress.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="ew")
            self.progress.set(0)

            self.lbl_prog = ctk.CTkLabel(frame_exec, text="0/0")
            self.lbl_prog.grid(row=1, column=0, padx=12, pady=(2, 6), sticky="e")

            ctk.CTkLabel(frame_exec, text="Log do processamento", font=ctk.CTkFont(size=14, weight="bold")).grid(
                row=2, column=0, padx=12, pady=(4, 4), sticky="nw"
            )

            self.txt_log = ctk.CTkTextbox(frame_exec)
            self.txt_log.grid(row=2, column=0, padx=12, pady=(30, 12), sticky="nsew")
            self.txt_log.configure(state="disabled")

            footer = ctk.CTkLabel(self, text=FOOTER_UI, font=ctk.CTkFont(size=11))
            footer.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="e")

        def _set_entry_readonly(self, entry, value: str):
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, value)
            entry.configure(state="readonly")

        def _selecionar_pasta(self):
            p = filedialog.askdirectory(title="Selecione a pasta com PDFs")

            if p:
                self.origem_pdf = Path(p)
                self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                self._log(f"Pasta selecionada: {self.origem_pdf}")

        def _selecionar_pdf(self):
            p = filedialog.askopenfilename(
                title="Selecione um PDF",
                filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")],
            )

            if p:
                self.origem_pdf = Path(p)
                self._set_entry_readonly(self.ent_origem, str(self.origem_pdf))
                self._log(f"PDF selecionado: {self.origem_pdf}")

        def _selecionar_base(self):
            p = filedialog.askopenfilename(
                title="Selecione a base",
                filetypes=[("Bases suportadas", "*.csv *.xls *.xlsx *.xlsm *.xlsb"), ("Todos", "*.*")],
            )

            if p:
                self.base_path = Path(p)
                self._set_entry_readonly(self.ent_base, str(self.base_path))
                self._log(f"Base selecionada: {self.base_path}")

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

        def _texto_base(self) -> str:
            texto = self.txt_base.get("1.0", "end").strip()
            exemplo = (
                "CPF;PIS;Matrícula;Nome;Email Alternativo\n"
                "12345678909;12345678901;123456;Nome e Sobrenome;email@exemplo.com"
            )

            if texto == exemplo:
                return ""

            return texto

        def _ao_mudar_tipo_documento(self, escolha: str):
            self._salvar_modelo_atual(silencioso=True)
            self.tipo_documento_atual = escolha
            self._carregar_template_na_tela()

        def _carregar_template_na_tela(self):
            template = self.templates_email.get(
                self.tipo_documento_atual,
                TEMPLATES_EMAIL["Informe de Rendimento"],
            )

            self.ent_assunto.delete(0, "end")
            self.ent_assunto.insert(0, template["assunto"])
            self.txt_corpo.delete("1.0", "end")
            self.txt_corpo.insert("1.0", template["corpo"])

        def _salvar_modelo_atual(self, silencioso: bool = False):
            assunto = self.ent_assunto.get().strip()
            corpo = self.txt_corpo.get("1.0", "end").strip()

            self.templates_email[self.tipo_documento_atual] = {
                "assunto": assunto,
                "corpo": corpo,
            }

            if not silencioso:
                self._log(f"Modelo salvo em memória para o tipo: {self.tipo_documento_atual}")
                messagebox.showinfo("Modelo salvo", f"Modelo atualizado para: {self.tipo_documento_atual}")

        def _restaurar_modelo_padrao(self):
            self.templates_email[self.tipo_documento_atual] = json.loads(
                json.dumps(TEMPLATES_EMAIL[self.tipo_documento_atual])
            )
            self._carregar_template_na_tela()
            self._log(f"Modelo padrão restaurado para: {self.tipo_documento_atual}")

        def _validar_base_colada(self):
            try:
                texto = self._texto_base()
                eh_ponto = self.tipo_documento_atual == "Ponto"

                if eh_ponto:
                    base = obter_base_pis(self.base_path, texto)
                    self._log(f"Base validada com sucesso. Total de PIS carregados: {len(base)}")
                    messagebox.showinfo("Base validada", f"Base validada com sucesso.\n\nTotal de PIS: {len(base)}")
                else:
                    base = obter_base(self.base_path, texto)
                    self._log(f"Base validada com sucesso. Total de CPFs carregados: {len(base)}")
                    messagebox.showinfo("Base validada", f"Base validada com sucesso.\n\nTotal de CPFs: {len(base)}")

            except Exception as e:
                self._log(f"Erro ao validar base: {e}")
                messagebox.showerror("Erro na base", str(e))

        def _alternar_botoes(self, habilitar: bool):
            estado = "normal" if habilitar else "disabled"
            self.btn_localizar.configure(state=estado)
            self.btn_iniciar.configure(state=estado)
            self.btn_validar_base.configure(state=estado)
            self.btn_reiniciar.configure(state=estado)
            self.btn_template.configure(state=estado)

        def _reiniciar(self):
            self.origem_pdf = None
            self.base_path = None
            self._set_entry_readonly(self.ent_origem, "")
            self._set_entry_readonly(self.ent_base, "")
            self.txt_base.delete("1.0", "end")
            self.txt_base.insert(
                "1.0",
                "CPF;PIS;Matrícula;Nome;Email Alternativo\n"
                "12345678909;12345678901;123456;Nome e Sobrenome;email@exemplo.com\n"
            )
            if hasattr(self, "ent_cc"):
                self.ent_cc.delete(0, "end")

            self._limpar_log()
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")
            self.cmb_tipo_documento.set("Informe de Rendimento")
            self.tipo_documento_atual = "Informe de Rendimento"
            self.templates_email = json.loads(json.dumps(TEMPLATES_EMAIL))
            self._carregar_template_na_tela()
            self._log("Tela reiniciada. Pronto para novo processamento.")

        def _gerar_template_excel(self):
            try:
                caminho = gerar_template_excel()
                self._log(f"Template Excel gerado: {caminho}")
                messagebox.showinfo("Template criado", f"O template foi criado com sucesso.\n\n{caminho}")
            except Exception as e:
                messagebox.showerror("Erro ao gerar template", str(e))

        def _executar(self, criar_rascunho: bool):
            if not self.origem_pdf:
                self._log("Erro: selecione uma pasta ou um PDF antes de iniciar.")
                messagebox.showerror("Erro", "Selecione uma pasta ou um PDF antes de iniciar.")
                return

            if criar_rascunho:
                texto = self._texto_base()

                if not self.base_path and not texto:
                    self._log("Erro: informe uma base por arquivo ou cole os dados na tela.")
                    messagebox.showerror(
                        "Erro",
                        "Para criar rascunhos, informe uma base por arquivo ou cole os dados diretamente na tela.",
                    )
                    return
            else:
                texto = self._texto_base()

            assunto_template = self.ent_assunto.get().strip()
            corpo_template = self.txt_corpo.get("1.0", "end").strip()
            cc_email = self.ent_cc.get().strip() if hasattr(self, "ent_cc") else ""

            self._alternar_botoes(False)
            self.progress.set(0)
            self.lbl_prog.configure(text="0/0")

            eh_ponto = self.tipo_documento_atual == "Ponto"

            if eh_ponto:
                modo = "Ponto por PIS sem senha" if criar_rascunho else "Encontrar PIS no(s) PDF(s)"
            else:
                modo = "Proteger + Criar Rascunho" if criar_rascunho else "Encontrar CPF(s)"

            self._log("")
            self._log("Iniciando processamento.")
            self._log(f"Modo: {modo}")
            self._log(f"Tipo documento: {self.tipo_documento_atual}")
            self._log(f"Origem: {self.origem_pdf}")

            if criar_rascunho:
                self._log("Base usada: base colada na tela" if texto else f"Base usada: {self.base_path}")

            if cc_email:
                self._log(f"CC/Cópia informado: {cc_email}")

            if eh_ponto:
                self._log("Regra aplicada: localizar por PIS e anexar PDF original sem senha.")
            else:
                self._log("Regra aplicada: localizar por CPF, proteger PDF com senha e criar rascunho.")

            self._log("")

            def worker():
                try:
                    if eh_ponto:
                        resultados, log_path = processar_arquivos_ponto_por_pis(
                            origem_pdf=self.origem_pdf,
                            base_path=self.base_path,
                            base_texto=texto,
                            criar_rascunho=criar_rascunho,
                            on_progress=lambda a, t: self.after(0, self._on_progress, a, t),
                            on_log=lambda m: self.after(0, self._log, m),
                            subject_template=assunto_template,
                            body_template=corpo_template,
                            cc_email=cc_email,
                        )
                    else:
                        resultados, log_path = processar_arquivos(
                            origem_pdf=self.origem_pdf,
                            base_path=self.base_path,
                            base_texto=texto,
                            criar_rascunho=criar_rascunho,
                            on_progress=lambda a, t: self.after(0, self._on_progress, a, t),
                            on_log=lambda m: self.after(0, self._log, m),
                            subject_template=assunto_template,
                            body_template=corpo_template,
                            cc_email=cc_email,
                        )

                    resumo = resumir_resultados(resultados, modo)
                    self.after(0, self._log, "")
                    self.after(0, self._log, resumo)
                    self.after(0, self._log, f"Log salvo em: {log_path}")
                    self.after(0, messagebox.showinfo, "Processo concluído", resumo + f"\n\nLog: {log_path}")

                except Exception as e:
                    msg = f"Ocorreu um erro: {e}\n\n{traceback.format_exc()}"
                    self.after(0, self._log, msg)
                    self.after(0, messagebox.showerror, "Erro", msg)

                finally:
                    self.after(0, self._alternar_botoes, True)

            threading.Thread(target=worker, daemon=True).start()

        def _iniciar_localizacao(self):
            self._executar(criar_rascunho=False)

        def _iniciar_rascunho(self):
            self._executar(criar_rascunho=True)


def main():
    if not CTK_AVAILABLE:
        raise RuntimeError("customtkinter não está instalado. Instale com: pip install customtkinter")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

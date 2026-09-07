from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

import pdfplumber
import xlsxwriter


# ============================================================
# CONFIGURAÇÃO VISUAL
# ============================================================
SONOVA_BLUE = "#0083CA"
LIGHT_BLUE = "#EAF5FB"
WHITE = "#FFFFFF"
DARK = "#333333"


# ============================================================
# MODELOS
# ============================================================
@dataclass
class DadosTrabalhador:
    trabalhador: str = ""
    cpf: str = ""
    matricula: str = ""
    data_admissao: datetime | None = None
    data_opcao_fgts: datetime | None = None
    data_desligamento: datetime | None = None
    empregador: str = ""
    local_trabalho: str = ""


@dataclass
class MovimentoFGTS:
    competencia: datetime | None
    origem: str
    remuneracao: float | str | None
    remuneracao_13: float | None
    aliquota: float | None
    fgts: float | None
    fgts_atualizado: float | None


@dataclass
class DadosRescisao:
    motivo: str = ""
    base_rescisoria: float | None = None
    percentual_multa: float | None = None
    indenizacao_compensatoria: float | None = None


# ============================================================
# FUNÇÕES DE CONVERSÃO
# ============================================================
def limpar_texto(valor: Any) -> str:
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def numero_br(valor: Any) -> float | None:
    """Converte 'R$ 1.234,56' ou '1.234,56' para float."""
    texto = limpar_texto(valor)
    if not texto:
        return None

    texto = texto.replace("R$", "").replace("%", "").replace(" ", "")
    texto = texto.replace(".", "").replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return None


def data_br(valor: Any) -> datetime | None:
    texto = limpar_texto(valor)
    if not texto:
        return None
    try:
        return datetime.strptime(texto, "%d/%m/%Y")
    except ValueError:
        return None


def competencia_br(valor: Any) -> datetime | None:
    texto = limpar_texto(valor)
    if not re.fullmatch(r"\d{2}/\d{4}", texto):
        return None
    try:
        return datetime.strptime("01/" + texto, "%d/%m/%Y")
    except ValueError:
        return None


def primeiro_nao_vazio(valores: Iterable[Any]) -> str:
    for valor in valores:
        texto = limpar_texto(valor)
        if texto:
            return texto
    return ""


def juntar_celulas(linha: list[Any]) -> list[str]:
    return [limpar_texto(x) for x in linha]


# ============================================================
# IDENTIFICAÇÃO DOS CAMPOS DO CABEÇALHO
# ============================================================
def extrair_dados_trabalhador(tabela: list[list[Any]]) -> DadosTrabalhador:
    """
    A ficha do FGTS geralmente traz duas linhas iniciais:
    1) cabeçalhos cadastrais
    2) valores cadastrais

    O pdfplumber pode inserir colunas vazias entre os campos, por isso
    fazemos o pareamento pelo nome do cabeçalho em vez de assumir índice fixo.
    """
    if len(tabela) < 2:
        return DadosTrabalhador()

    cab = juntar_celulas(tabela[0])
    val = juntar_celulas(tabela[1])

    # Garante mesmo tamanho
    if len(val) < len(cab):
        val += [""] * (len(cab) - len(val))

    mapa: dict[str, str] = {}
    ultimo_header = ""

    for i, header in enumerate(cab):
        if header:
            ultimo_header = header
        if ultimo_header and i < len(val) and val[i]:
            # Em PDFs com células mescladas, o valor pode cair numa coluna
            # imediatamente posterior ao nome do cabeçalho.
            if ultimo_header not in mapa:
                mapa[ultimo_header] = val[i]

    def achar(*nomes: str) -> str:
        for chave, valor in mapa.items():
            chave_norm = chave.lower()
            if any(nome.lower() in chave_norm for nome in nomes):
                return valor
        return ""

    return DadosTrabalhador(
        trabalhador=achar("trabalhador"),
        cpf=achar("cpf"),
        matricula=achar("matrícula", "matricula"),
        data_admissao=data_br(achar("data de admissão", "data de admissao")),
        data_opcao_fgts=data_br(achar("data de opção fgts", "data de opcao fgts")),
        data_desligamento=data_br(achar("data de desligamento")),
        empregador=achar("empregador"),
        local_trabalho=achar("local de trabalho"),
    )


# ============================================================
# EXTRAÇÃO DAS LINHAS MENSAIS
# ============================================================
def linha_e_movimento(linha: list[Any]) -> bool:
    if not linha:
        return False
    primeira = limpar_texto(linha[0])
    return bool(re.fullmatch(r"\d{2}/\d{4}", primeira))


def localizar_valores_movimento(linha: list[Any]) -> MovimentoFGTS | None:
    """
    Faz o parser sem depender rigidamente da quantidade de colunas,
    porque o PDF pode quebrar 'R$' e valores em células distintas.

    Estratégia:
    - competência = primeira célula
    - origem = primeira célula textual após competência
    - 'Afastamento' pode aparecer em Remuneração
    - números monetários são capturados na ordem visual
    - alíquota costuma ser 8
    """
    celulas = juntar_celulas(linha)
    if not celulas:
        return None

    competencia = competencia_br(celulas[0])
    if not competencia:
        return None

    origem = ""
    for item in celulas[1:]:
        if item and not re.search(r"\d", item) and "R$" not in item and item.lower() != "afastamento":
            origem = item
            break

    afastamento = any(item.lower() == "afastamento" for item in celulas)

    # Captura valores monetários explicitamente marcados com R$
    monetarios: list[float] = []
    for item in celulas:
        if "R$" in item:
            n = numero_br(item)
            if n is not None:
                monetarios.append(n)

    # Também captura valores como 227,38 que aparecem sem R$ na coluna FGTS Atualizado.
    numeros_sem_rs: list[float] = []
    for i, item in enumerate(celulas[1:], start=1):
        if not item or "R$" in item or item.lower() == "afastamento":
            continue
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", item):
            n = numero_br(item)
            if n is not None:
                numeros_sem_rs.append(n)

    # Alíquota: normalmente 8. Evitamos interpretar partes de moeda como alíquota.
    aliquota = None
    for item in celulas[1:]:
        if item in {"8", "8,00", "8.00"}:
            aliquota = 8.0
            break

    remuneracao: float | str | None = "Afastamento" if afastamento else None
    remuneracao_13: float | None = None
    fgts: float | None = None
    fgts_atualizado: float | None = None

    if afastamento:
        # Em afastamentos, podem existir 13º e/ou FGTS. A ordem visual pode variar.
        if len(monetarios) >= 2:
            remuneracao_13 = monetarios[0]
            fgts = monetarios[-1]
        elif len(monetarios) == 1:
            fgts = monetarios[0]
    else:
        if len(monetarios) >= 1:
            remuneracao = monetarios[0]
        if len(monetarios) >= 3:
            remuneracao_13 = monetarios[1]
            fgts = monetarios[2]
        elif len(monetarios) >= 2:
            fgts = monetarios[1]

    # FGTS atualizado costuma vir sem 'R$' e ser o último decimal sem R$.
    if numeros_sem_rs:
        fgts_atualizado = numeros_sem_rs[-1]

    return MovimentoFGTS(
        competencia=competencia,
        origem=origem,
        remuneracao=remuneracao,
        remuneracao_13=remuneracao_13,
        aliquota=aliquota,
        fgts=fgts,
        fgts_atualizado=fgts_atualizado,
    )


def extrair_movimentos(tabelas: list[list[list[Any]]]) -> list[MovimentoFGTS]:
    movimentos: list[MovimentoFGTS] = []

    for tabela in tabelas:
        for linha in tabela:
            if linha_e_movimento(linha):
                mov = localizar_valores_movimento(linha)
                if mov:
                    movimentos.append(mov)

    # Remove eventuais duplicidades por competência/origem/valores
    unicos: list[MovimentoFGTS] = []
    vistos = set()
    for m in movimentos:
        chave = (
            m.competencia,
            m.origem,
            m.remuneracao,
            m.remuneracao_13,
            m.aliquota,
            m.fgts,
            m.fgts_atualizado,
        )
        if chave not in vistos:
            vistos.add(chave)
            unicos.append(m)

    return sorted(unicos, key=lambda x: x.competencia or datetime.min)


# ============================================================
# EXTRAÇÃO DA RESCISÃO
# ============================================================
def extrair_rescisao(tabelas: list[list[list[Any]]]) -> DadosRescisao:
    for tabela in tabelas:
        for i, linha in enumerate(tabela):
            texto = " | ".join(juntar_celulas(linha)).lower()
            if "motivo do desligamento" in texto:
                if i + 1 >= len(tabela):
                    continue

                prox = juntar_celulas(tabela[i + 1])
                motivo = primeiro_nao_vazio(prox)

                monetarios = [numero_br(x) for x in prox if "R$" in x]
                monetarios = [x for x in monetarios if x is not None]

                percentual = None
                for item in prox:
                    if "%" in item:
                        n = numero_br(item)
                        if n is not None:
                            percentual = n / 100
                            break

                base = monetarios[0] if len(monetarios) >= 1 else None
                indenizacao = monetarios[-1] if len(monetarios) >= 2 else None

                return DadosRescisao(
                    motivo=motivo,
                    base_rescisoria=base,
                    percentual_multa=percentual,
                    indenizacao_compensatoria=indenizacao,
                )

    return DadosRescisao()


# ============================================================
# LEITURA DO PDF
# ============================================================
def ler_pdf(pdf_path: Path) -> tuple[DadosTrabalhador, list[MovimentoFGTS], DadosRescisao]:
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF não encontrado: {pdf_path}")

    tabelas: list[list[list[Any]]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            extraidas = pagina.extract_tables()
            if extraidas:
                tabelas.extend(extraidas)

    if not tabelas:
        raise RuntimeError(
            "Nenhuma tabela foi localizada no PDF. "
            "Este script funciona para fichas de remuneração/FGTS com texto pesquisável."
        )

    trabalhador = extrair_dados_trabalhador(tabelas[0])
    movimentos = extrair_movimentos(tabelas)
    rescisao = extrair_rescisao(tabelas)

    return trabalhador, movimentos, rescisao


# ============================================================
# GERAÇÃO DO EXCEL
# ============================================================
def gerar_excel(
    destino: Path,
    trabalhador: DadosTrabalhador,
    movimentos: list[MovimentoFGTS],
    rescisao: DadosRescisao,
    fonte_pdf: Path,
) -> None:
    destino.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(destino)
    workbook.set_properties({
        "title": "Ficha de Remuneração / FGTS",
        "subject": "Conversão de PDF para Excel",
        "author": "Igarapé Digital",
        "comments": f"Gerado a partir de {fonte_pdf.name}",
    })

    # Formatos
    fmt_titulo = workbook.add_format({
        "bold": True,
        "font_color": WHITE,
        "bg_color": SONOVA_BLUE,
        "align": "center",
        "valign": "vcenter",
        "font_size": 14,
    })
    fmt_header = workbook.add_format({
        "bold": True,
        "font_color": WHITE,
        "bg_color": SONOVA_BLUE,
        "align": "center",
        "valign": "vcenter",
        "border": 1,
    })
    fmt_rotulo = workbook.add_format({
        "bold": True,
        "bg_color": LIGHT_BLUE,
        "font_color": DARK,
        "border": 1,
        "valign": "top",
    })
    fmt_texto = workbook.add_format({"border": 1, "valign": "top"})
    fmt_data = workbook.add_format({"border": 1, "num_format": "dd/mm/yyyy"})
    fmt_comp = workbook.add_format({"border": 1, "num_format": "mm/yyyy"})
    fmt_moeda = workbook.add_format({"border": 1, "num_format": 'R$ #,##0.00'})
    fmt_numero = workbook.add_format({"border": 1, "num_format": "0.00"})
    fmt_percentual = workbook.add_format({"border": 1, "num_format": "0.00%"})

    # ------------------- ABA RESUMO -------------------
    ws = workbook.add_worksheet("Resumo")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 38)
    ws.set_column("B:B", 58)
    ws.set_row(0, 24)
    ws.merge_range("A1:B1", "Ficha de Remuneração / FGTS", fmt_titulo)

    ws.write("A3", "Dados do Trabalhador", fmt_header)
    ws.write("B3", "Valor", fmt_header)

    dados = [
        ("Trabalhador", trabalhador.trabalhador, "text"),
        ("CPF", trabalhador.cpf, "text"),
        ("Matrícula", trabalhador.matricula, "text"),
        ("Data de Admissão", trabalhador.data_admissao, "date"),
        ("Data de Opção FGTS", trabalhador.data_opcao_fgts, "date"),
        ("Data de Desligamento", trabalhador.data_desligamento, "date"),
        ("Empregador", trabalhador.empregador, "text"),
        ("Local de Trabalho", trabalhador.local_trabalho, "text"),
    ]

    linha = 3
    for rotulo, valor, tipo in dados:
        ws.write(linha, 0, rotulo, fmt_rotulo)
        if tipo == "date" and isinstance(valor, datetime):
            ws.write_datetime(linha, 1, valor, fmt_data)
        else:
            ws.write(linha, 1, valor or "", fmt_texto)
        linha += 1

    linha += 1
    ws.write(linha, 0, "Dados da Rescisão", fmt_header)
    ws.write(linha, 1, "Valor", fmt_header)
    linha += 1

    ws.write(linha, 0, "Motivo do Desligamento", fmt_rotulo)
    ws.write(linha, 1, rescisao.motivo, fmt_texto)
    linha += 1

    ws.write(linha, 0, "Valor da base para fins rescisórios", fmt_rotulo)
    if rescisao.base_rescisoria is not None:
        ws.write_number(linha, 1, rescisao.base_rescisoria, fmt_moeda)
    else:
        ws.write_blank(linha, 1, None, fmt_moeda)
    linha += 1

    ws.write(linha, 0, "Percentual da Multa", fmt_rotulo)
    if rescisao.percentual_multa is not None:
        ws.write_number(linha, 1, rescisao.percentual_multa, fmt_percentual)
    else:
        ws.write_blank(linha, 1, None, fmt_percentual)
    linha += 1

    ws.write(linha, 0, "Indenização Compensatória", fmt_rotulo)
    if rescisao.indenizacao_compensatoria is not None:
        ws.write_number(linha, 1, rescisao.indenizacao_compensatoria, fmt_moeda)
    else:
        ws.write_blank(linha, 1, None, fmt_moeda)
    linha += 2

    ws.write(linha, 0, "Fonte", fmt_rotulo)
    ws.write(linha, 1, fonte_pdf.name, fmt_texto)

    # ------------------- ABA MOVIMENTAÇÃO -------------------
    mov = workbook.add_worksheet("Movimentação FGTS")
    mov.hide_gridlines(2)
    mov.freeze_panes(1, 0)

    colunas = [
        ("Competência", 14),
        ("Origem", 18),
        ("Remuneração", 20),
        ("Remuneração 13º", 20),
        ("Alíquota (%)", 14),
        ("FGTS", 16),
        ("FGTS Atualizado (R$)", 22),
    ]

    for col, (titulo, largura) in enumerate(colunas):
        mov.write(0, col, titulo, fmt_header)
        mov.set_column(col, col, largura)

    for row, m in enumerate(movimentos, start=1):
        if m.competencia:
            mov.write_datetime(row, 0, m.competencia, fmt_comp)
        else:
            mov.write_blank(row, 0, None, fmt_comp)

        mov.write(row, 1, m.origem, fmt_texto)

        if isinstance(m.remuneracao, (int, float)):
            mov.write_number(row, 2, float(m.remuneracao), fmt_moeda)
        elif isinstance(m.remuneracao, str):
            mov.write(row, 2, m.remuneracao, fmt_texto)
        else:
            mov.write_blank(row, 2, None, fmt_moeda)

        if m.remuneracao_13 is not None:
            mov.write_number(row, 3, m.remuneracao_13, fmt_moeda)
        else:
            mov.write_blank(row, 3, None, fmt_moeda)

        if m.aliquota is not None:
            mov.write_number(row, 4, m.aliquota, fmt_numero)
        else:
            mov.write_blank(row, 4, None, fmt_numero)

        if m.fgts is not None:
            mov.write_number(row, 5, m.fgts, fmt_moeda)
        else:
            mov.write_blank(row, 5, None, fmt_moeda)

        if m.fgts_atualizado is not None:
            mov.write_number(row, 6, m.fgts_atualizado, fmt_moeda)
        else:
            mov.write_blank(row, 6, None, fmt_moeda)

    # Cria tabela Excel com filtro automático
    if movimentos:
        mov.add_table(
            0,
            0,
            len(movimentos),
            6,
            {
                "name": "MovimentacaoFGTS",
                "style": "Table Style Medium 2",
                "columns": [{"header": titulo} for titulo, _ in colunas],
            },
        )

    workbook.close()


# ============================================================
# INTERFACE / EXECUÇÃO
# ============================================================
def escolher_pdf_gui() -> Path | None:
    """Abre seletor de arquivo caso nenhum PDF seja informado no terminal."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        arquivo = filedialog.askopenfilename(
            title="Selecione a ficha de remuneração em PDF",
            filetypes=[("Arquivos PDF", "*.pdf")],
        )
        root.destroy()
        return Path(arquivo) if arquivo else None
    except Exception:
        return None


def criar_nome_saida(pdf_path: Path) -> Path:
    return pdf_path.with_name(f"{pdf_path.stem}_convertido.xlsx")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Converte ficha de remuneração/FGTS em PDF para Excel."
    )
    parser.add_argument(
        "pdf",
        nargs="?",
        help="Caminho do PDF. Se omitido, abre uma janela para seleção.",
    )
    parser.add_argument(
        "-o",
        "--saida",
        help="Caminho do Excel de saída. Se omitido, salva ao lado do PDF.",
    )

    args = parser.parse_args()

    pdf_path = Path(args.pdf).expanduser().resolve() if args.pdf else escolher_pdf_gui()
    if not pdf_path:
        print("Nenhum PDF selecionado.")
        return

    if not pdf_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {pdf_path}")

    destino = (
        Path(args.saida).expanduser().resolve()
        if args.saida
        else criar_nome_saida(pdf_path)
    )

    print(f"Lendo: {pdf_path}")
    trabalhador, movimentos, rescisao = ler_pdf(pdf_path)

    if not movimentos:
        print("ATENÇÃO: não foram localizadas competências mensais.")

    gerar_excel(destino, trabalhador, movimentos, rescisao, pdf_path)

    print("\nConversão concluída.")
    print(f"Trabalhador: {trabalhador.trabalhador or 'não identificado'}")
    print(f"Competências extraídas: {len(movimentos)}")
    print(f"Excel gerado: {destino}")


if __name__ == "__main__":
    main()

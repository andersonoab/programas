# ============================================================
# App: Renomeador Curto de PDFs PGR / PCMSO / LTCAT
# Autor: Anderson Marinho | Igarapé Digital
# Objetivo:
#   Selecionar uma pasta e renomear todos os PDFs com nomes longos
#   para padrão curto, aproveitando tipo documental e CNPJ/CPF.
#
# Exemplo:
#   LTCAT COMPLETO - AUDITIV COMERCIO E SERVICOS ... 23.689.7740001-77.pdf
#   vira:
#   LTCAT_23689774000177.pdf
# ============================================================

import os
import re
import shutil
import unicodedata
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import customtkinter as ctk


# ============================================================
# TEMA CUSTOMERTHINKER
# ============================================================

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

COR_NAVY = "#121C4E"
COR_AZUL = "#0083CA"
COR_AZUL_ESCURO = "#003C64"
COR_CINZA_BG = "#F4F6F8"
COR_CARD = "#FFFFFF"
COR_TEXTO = "#333333"
COR_SUCESSO = "#005A64"
COR_ALERTA = "#8C321E"


# ============================================================
# FUNÇÕES DE APOIO
# ============================================================

def remover_acentos(texto: str) -> str:
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def limpar_texto(texto: str) -> str:
    texto = remover_acentos(texto or "")
    texto = texto.upper()
    texto = re.sub(r"[^A-Z0-9\s._-]", " ", texto)
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def somente_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def identificar_tipo_documento(nome: str) -> str:
    """
    Identifica o tipo principal do documento.
    """
    nome_limpo = limpar_texto(nome)

    if "LTCAT" in nome_limpo:
        return "LTCAT"

    if "PCMSO" in nome_limpo:
        return "PCMSO"

    if "PGR" in nome_limpo:
        return "PGR"

    if "PPRA" in nome_limpo:
        return "PPRA"

    if "ASO" in nome_limpo:
        return "ASO"

    if "LAUDO" in nome_limpo:
        return "LAUDO"

    return "DOCUMENTO"


def capturar_cnpj_ou_cpf(nome: str) -> str:
    """
    Captura CNPJ ou CPF dentro do nome.
    Primeiro tenta CNPJ com 14 dígitos.
    Depois tenta CPF com 11 dígitos.
    """
    texto = nome or ""

    # Captura CNPJ com máscara clássica: 23.689.774/0001-77
    padrao_cnpj_mascara = r"\d{2}\.?\d{3}\.?\d{3}[\/.-]?\d{4}-?\d{2}"
    encontrados = re.findall(padrao_cnpj_mascara, texto)

    for item in encontrados:
        digitos = somente_digitos(item)
        if len(digitos) == 14:
            return digitos

    # Captura qualquer sequência numérica possível
    blocos = re.findall(r"\d[\d.\-_/ ]{8,25}\d", texto)

    for bloco in blocos:
        digitos = somente_digitos(bloco)

        if len(digitos) == 14:
            return digitos

    for bloco in blocos:
        digitos = somente_digitos(bloco)

        if len(digitos) == 11:
            return digitos

    return ""


def capturar_empresa_curta(nome: str) -> str:
    """
    Tenta aproveitar uma palavra útil do nome da empresa.
    Exemplo:
    AUDITIV COMERCIO E SERVICOS...
    Retorna AUDITIV.
    """
    texto = limpar_texto(nome)

    palavras_ignoradas = {
        "LTCAT", "PCMSO", "PGR", "PPRA", "ASO", "LAUDO",
        "COMPLETO", "COMPLETA", "DOCUMENTO", "DOCUMENTOS",
        "COMERCIO", "SERVICOS", "SERVICO", "PRODUTOS",
        "HOSPITALARES", "LTDA", "E", "DE", "DA", "DO",
        "DAS", "DOS", "PARA", "COM"
    }

    partes = re.split(r"[\s._-]+", texto)

    for parte in partes:
        parte = parte.strip()

        if not parte:
            continue

        if parte in palavras_ignoradas:
            continue

        if len(parte) <= 2:
            continue

        if parte.isdigit():
            continue

        return parte[:18]

    return ""


def gerar_nome_curto(caminho_pdf: Path, usar_empresa: bool = False) -> str:
    nome_original = caminho_pdf.name

    tipo = identificar_tipo_documento(nome_original)
    documento = capturar_cnpj_ou_cpf(nome_original)
    empresa = capturar_empresa_curta(nome_original)

    partes = [tipo]

    if usar_empresa and empresa:
        partes.append(empresa)

    if documento:
        partes.append(documento)

    if not documento:
        data_ref = datetime.now().strftime("%Y%m%d")
        partes.append(data_ref)

    novo_nome = "_".join(partes)
    novo_nome = re.sub(r"_+", "_", novo_nome).strip("_")

    return f"{novo_nome}.pdf"


def gerar_nome_unico(caminho_destino: Path) -> Path:
    """
    Evita sobrescrever arquivos.
    """
    if not caminho_destino.exists():
        return caminho_destino

    pasta = caminho_destino.parent
    stem = caminho_destino.stem
    suffix = caminho_destino.suffix

    contador = 2

    while True:
        novo = pasta / f"{stem}_{contador:03d}{suffix}"
        if not novo.exists():
            return novo
        contador += 1


# ============================================================
# APP
# ============================================================

class RenomeadorPDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Renomeador Curto de PDFs - PGR PCMSO LTCAT")
        self.geometry("1180x720")
        self.minsize(1050, 650)
        self.configure(fg_color=COR_CINZA_BG)

        self.pasta = None
        self.dados = []

        self.var_usar_empresa = tk.BooleanVar(value=False)
        self.var_modo = tk.StringVar(value="renomear")
        self.var_recursivo = tk.BooleanVar(value=False)

        self.criar_tela()

    def criar_tela(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(self, fg_color=COR_NAVY, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")

        titulo = ctk.CTkLabel(
            header,
            text="Renomeador Curto de PDFs",
            font=("Arial", 24, "bold"),
            text_color="white"
        )
        titulo.pack(anchor="w", padx=20, pady=(16, 2))

        subtitulo = ctk.CTkLabel(
            header,
            text="Reduz nomes longos de PDFs usando tipo documental e CNPJ/CPF",
            font=("Arial", 13),
            text_color="white"
        )
        subtitulo.pack(anchor="w", padx=20, pady=(0, 16))

        painel = ctk.CTkFrame(self, fg_color=COR_CARD, corner_radius=12)
        painel.grid(row=1, column=0, padx=16, pady=14, sticky="ew")
        painel.grid_columnconfigure(1, weight=1)

        btn_pasta = ctk.CTkButton(
            painel,
            text="Selecionar pasta",
            command=self.selecionar_pasta,
            fg_color=COR_AZUL,
            hover_color=COR_AZUL_ESCURO,
            height=36
        )
        btn_pasta.grid(row=0, column=0, padx=12, pady=12)

        self.lbl_pasta = ctk.CTkLabel(
            painel,
            text="Nenhuma pasta selecionada",
            text_color=COR_TEXTO,
            anchor="w"
        )
        self.lbl_pasta.grid(row=0, column=1, padx=8, pady=12, sticky="ew")

        btn_analisar = ctk.CTkButton(
            painel,
            text="Analisar",
            command=self.analisar,
            fg_color=COR_AZUL_ESCURO,
            hover_color=COR_NAVY,
            height=36
        )
        btn_analisar.grid(row=0, column=2, padx=8, pady=12)

        btn_executar = ctk.CTkButton(
            painel,
            text="Executar",
            command=self.executar,
            fg_color=COR_SUCESSO,
            hover_color=COR_NAVY,
            height=36
        )
        btn_executar.grid(row=0, column=3, padx=12, pady=12)

        opcoes = ctk.CTkFrame(self, fg_color=COR_CARD, corner_radius=12)
        opcoes.grid(row=2, column=0, padx=16, pady=(0, 14), sticky="nsew")
        opcoes.grid_columnconfigure(0, weight=1)
        opcoes.grid_rowconfigure(1, weight=1)

        linha_opcoes = ctk.CTkFrame(opcoes, fg_color=COR_CARD)
        linha_opcoes.grid(row=0, column=0, padx=12, pady=10, sticky="ew")

        check_empresa = ctk.CTkCheckBox(
            linha_opcoes,
            text="Incluir empresa curta no nome",
            variable=self.var_usar_empresa,
            text_color=COR_TEXTO
        )
        check_empresa.grid(row=0, column=0, padx=8, pady=8, sticky="w")

        check_recursivo = ctk.CTkCheckBox(
            linha_opcoes,
            text="Buscar também em subpastas",
            variable=self.var_recursivo,
            text_color=COR_TEXTO
        )
        check_recursivo.grid(row=0, column=1, padx=18, pady=8, sticky="w")

        ctk.CTkLabel(
            linha_opcoes,
            text="Modo:",
            text_color=COR_TEXTO
        ).grid(row=0, column=2, padx=(18, 6), pady=8)

        menu_modo = ctk.CTkOptionMenu(
            linha_opcoes,
            values=["renomear", "copiar"],
            variable=self.var_modo,
            fg_color=COR_AZUL,
            button_color=COR_AZUL_ESCURO,
            button_hover_color=COR_NAVY,
            width=120
        )
        menu_modo.grid(row=0, column=3, padx=6, pady=8)

        self.lbl_status = ctk.CTkLabel(
            linha_opcoes,
            text="Pronto",
            text_color=COR_TEXTO
        )
        self.lbl_status.grid(row=0, column=4, padx=20, pady=8, sticky="w")

        tabela_frame = ctk.CTkFrame(opcoes, fg_color=COR_CARD)
        tabela_frame.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")
        tabela_frame.grid_columnconfigure(0, weight=1)
        tabela_frame.grid_rowconfigure(0, weight=1)

        colunas = ("pasta", "original", "novo", "status")

        self.tree = ttk.Treeview(
            tabela_frame,
            columns=colunas,
            show="headings"
        )

        self.tree.heading("pasta", text="Pasta")
        self.tree.heading("original", text="Nome original")
        self.tree.heading("novo", text="Novo nome")
        self.tree.heading("status", text="Status")

        self.tree.column("pasta", width=260, anchor="w")
        self.tree.column("original", width=430, anchor="w")
        self.tree.column("novo", width=270, anchor="w")
        self.tree.column("status", width=140, anchor="center")

        scroll_y = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(tabela_frame, orient="horizontal", command=self.tree.xview)

        self.tree.configure(
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set
        )

        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        rodape = ctk.CTkFrame(self, fg_color=COR_NAVY, corner_radius=0)
        rodape.grid(row=3, column=0, sticky="ew")

        ctk.CTkLabel(
            rodape,
            text="Anderson Marinho | Igarapé Digital",
            font=("Arial", 11),
            text_color="white"
        ).pack(side="right", padx=16, pady=6)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta dos PDFs")

        if not pasta:
            return

        self.pasta = Path(pasta)
        self.lbl_pasta.configure(text=str(self.pasta))
        self.lbl_status.configure(text="Pasta selecionada")
        self.limpar_tabela()

    def limpar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.dados = []

    def listar_pdfs(self):
        if not self.pasta:
            return []

        if self.var_recursivo.get():
            return sorted(self.pasta.rglob("*.pdf"))

        return sorted(self.pasta.glob("*.pdf"))

    def analisar(self):
        if not self.pasta:
            messagebox.showwarning("Atenção", "Selecione uma pasta primeiro.")
            return

        self.limpar_tabela()

        pdfs = self.listar_pdfs()

        if not pdfs:
            messagebox.showinfo("Resultado", "Nenhum PDF encontrado.")
            return

        usar_empresa = self.var_usar_empresa.get()

        for pdf in pdfs:
            novo_nome = gerar_nome_curto(pdf, usar_empresa=usar_empresa)

            status = "Pendente"
            if pdf.name.lower() == novo_nome.lower():
                status = "Já está curto"

            registro = {
                "pasta": str(pdf.parent),
                "caminho_original": str(pdf),
                "nome_original": pdf.name,
                "novo_nome": novo_nome,
                "status": status,
                "caminho_final": ""
            }

            self.dados.append(registro)

            self.tree.insert(
                "",
                "end",
                values=(
                    str(pdf.parent),
                    pdf.name,
                    novo_nome,
                    status
                )
            )

        self.lbl_status.configure(text=f"{len(self.dados)} PDF(s) analisado(s)")

    def executar(self):
        if not self.dados:
            messagebox.showwarning("Atenção", "Clique em Analisar antes de executar.")
            return

        modo = self.var_modo.get()
        processados = 0
        ignorados = 0
        erros = 0

        pasta_copia = None

        if modo == "copiar":
            pasta_copia = self.pasta / "_PDFS_RENOMEADOS_CURTOS"
            pasta_copia.mkdir(exist_ok=True)

        for registro in self.dados:
            try:
                origem = Path(registro["caminho_original"])

                if not origem.exists():
                    registro["status"] = "Arquivo não encontrado"
                    erros += 1
                    continue

                if registro["status"] == "Já está curto":
                    ignorados += 1
                    continue

                if modo == "copiar":
                    destino = pasta_copia / registro["novo_nome"]
                    destino = gerar_nome_unico(destino)
                    shutil.copy2(origem, destino)
                    registro["status"] = "Copiado"

                else:
                    destino = origem.parent / registro["novo_nome"]
                    destino = gerar_nome_unico(destino)
                    origem.rename(destino)
                    registro["status"] = "Renomeado"

                registro["caminho_final"] = str(destino)
                processados += 1

            except Exception as e:
                registro["status"] = f"Erro: {e}"
                erros += 1

        self.atualizar_tabela()
        self.salvar_log()

        self.lbl_status.configure(
            text=f"Processados: {processados} | Ignorados: {ignorados} | Erros: {erros}"
        )

        messagebox.showinfo(
            "Finalizado",
            f"Processados: {processados}\nIgnorados: {ignorados}\nErros: {erros}"
        )

    def atualizar_tabela(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for registro in self.dados:
            self.tree.insert(
                "",
                "end",
                values=(
                    registro["pasta"],
                    registro["nome_original"],
                    registro["novo_nome"],
                    registro["status"]
                )
            )

    def salvar_log(self):
        if not self.pasta:
            return

        agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_log = self.pasta / f"log_renomeador_pdf_{agora}.xlsx"

        df = pd.DataFrame(self.dados)

        try:
            df.to_excel(caminho_log, index=False)
        except Exception as e:
            messagebox.showwarning(
                "Aviso",
                f"Processo executado, mas não consegui salvar o log Excel:\n{e}"
            )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    app = RenomeadorPDFApp()
    app.mainloop()

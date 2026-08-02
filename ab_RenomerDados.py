"""
Renomeador de Arquivos - Remocao de Prefixo (ex.: CPF_)
Interface desktop em CustomTkinter.

Requisitos:
    pip install customtkinter

Rodar:
    python renomeador_desktop.py
"""

import os
import re
from pathlib import Path

import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox

# ==================== TEMA ====================
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

AZUL = "#0083CA"
AZUL_ESCURO = "#003C64"
AZUL_HOVER = "#005A8A"
AZUL_CLARO = "#6EB4DC"
ZEBRA = "#F2F8FC"
VINHO = "#7D0041"
CINZA = "#646464"
TEXTO = "#333333"

# Biblioteca de padroes prontos (nome -> regex). Voltada a documentos de DP/RH.
# Todos aceitam a forma com ou sem pontuacao.
PADROES = {
    "CPF": r"\d{3}[.\s]?\d{3}[.\s]?\d{3}[-\s]?\d{2}",
    "PIS/PASEP/NIT/NIS": r"\d{3}[.\s]?\d{5}[.\s]?\d{2}[-\s]?\d{1}",
    "CNPJ": r"\d{2}[.\s]?\d{3}[.\s]?\d{3}[/\s]?\d{4}[-\s]?\d{2}",
    "RG (SP)": r"\d{2}[.\s]?\d{3}[.\s]?\d{3}[-\s]?[\dxX]",
    "CNH": r"\d{11}",
    "Cartao SUS (CNS)": r"\d{3}[.\s]?\d{4}[.\s]?\d{4}[.\s]?\d{4}",
    "Titulo de eleitor": r"\d{4}[.\s]?\d{4}[.\s]?\d{4}",
    "CEP": r"\d{5}[-\s]?\d{3}",
    "Telefone": r"\(?\d{2}\)?[\s-]?9?\d{4}[\s-]?\d{4}",
    "Data (dd/mm/aaaa)": r"\d{2}[/.\-]\d{2}[/.\-]\d{4}",
    "Matricula (4 a 8 digitos)": r"\d{4,8}",
    "Qualquer numero (sequencia de digitos)": r"\d+",
}


class AutoHideScrollbar(ctk.CTkScrollbar):
    """Scrollbar fina que se esconde quando todo o conteudo esta visivel."""

    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
        super().set(lo, hi)


class DetalheModal(ctk.CTkToplevel):
    """Modal com o detalhe completo de um registro (duplo-clique na linha)."""

    def __init__(self, master, registro):
        super().__init__(master)
        self.title("Detalhe do arquivo")
        self.geometry("560x340")
        self.resizable(False, False)
        self.grab_set()

        header = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=44)
        header.pack(fill="x")
        ctk.CTkLabel(
            header, text="Detalhe do arquivo",
            text_color="#FFFFFF", font=("Segoe UI", 15, "bold"),
        ).pack(side="left", padx=16, pady=8)

        corpo = ctk.CTkFrame(self, fg_color="#FFFFFF")
        corpo.pack(fill="both", expand=True, padx=0, pady=0)

        campos = [
            ("Nome original", registro["original"]),
            ("Novo nome", registro["novo"]),
            ("Status", registro["status"]),
            ("Motivo", registro["motivo"] or "-"),
            ("Caminho", str(registro["path"])),
        ]

        for i, (rotulo, valor) in enumerate(campos):
            ctk.CTkLabel(
                corpo, text=rotulo, text_color=CINZA,
                font=("Segoe UI", 11, "bold"), anchor="w",
            ).grid(row=i, column=0, sticky="w", padx=(18, 10), pady=8)
            ctk.CTkLabel(
                corpo, text=valor, text_color=TEXTO,
                font=("Segoe UI", 11), anchor="w", wraplength=360, justify="left",
            ).grid(row=i, column=1, sticky="w", padx=(0, 18), pady=8)

        corpo.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            self, text="Fechar", fg_color=AZUL, hover_color=AZUL_HOVER,
            command=self.destroy, width=120,
        ).pack(pady=(0, 16))


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Renomeador de Arquivos - Igarape Digital")
        self.geometry("980x640")
        self.minsize(820, 520)

        # Estado
        self.registros = []          # lista de dicts: original, novo, status, motivo, path
        self.patterns_ativos = []    # dicts: {nome, regex, compiled}
        self.patterns = []           # regex compilados usados na analise
        self.sort_ascending = {}     # coluna -> bool
        self.coluna_ativa = None

        self._montar_ttk_style()
        self._montar_header()
        self._montar_controles()
        self._montar_tabela()
        self._montar_rodape()

    # ---------- Estilos ----------
    def _montar_ttk_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#FFFFFF", foreground=TEXTO, fieldbackground="#FFFFFF",
            rowheight=28, borderwidth=0, font=("Segoe UI", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=AZUL, foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"), relief="flat", padding=6,
        )
        style.map("Treeview.Heading", background=[("active", AZUL_HOVER)])
        style.map(
            "Treeview",
            background=[("selected", AZUL_CLARO)],
            foreground=[("selected", AZUL_ESCURO)],
        )

    # ---------- Header ----------
    def _montar_header(self):
        header = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header, text="Renomeador de Arquivos",
            text_color="#FFFFFF", font=("Segoe UI", 18, "bold"),
        ).pack(side="left", padx=20)
        ctk.CTkLabel(
            header, text="Remocao de prefixo (ex.: CPF_)",
            text_color="#E6F3FB", font=("Segoe UI", 12),
        ).pack(side="left")

    # ---------- Controles ----------
    def _montar_controles(self):
        painel = ctk.CTkFrame(self, fg_color="#FFFFFF")
        painel.pack(fill="x", padx=16, pady=(14, 6))

        # Linha 1: pasta
        linha1 = ctk.CTkFrame(painel, fg_color="transparent")
        linha1.pack(fill="x", pady=4)
        ctk.CTkLabel(linha1, text="Pasta:", width=90, anchor="w",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left")
        self.entry_pasta = ctk.CTkEntry(linha1, placeholder_text="Selecione a pasta com os arquivos")
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(linha1, text="Procurar", width=110, fg_color=AZUL,
                      hover_color=AZUL_HOVER, command=self.escolher_pasta).pack(side="left")

        # Linha 2: prefixos + extensoes
        linha2 = ctk.CTkFrame(painel, fg_color="transparent")
        linha2.pack(fill="x", pady=4)
        ctk.CTkLabel(linha2, text="Prefixos:", width=90, anchor="w",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left")
        self.entry_prefixos = ctk.CTkEntry(linha2, width=220)
        self.entry_prefixos.insert(0, "CPF_, CPF")
        self.entry_prefixos.pack(side="left", padx=(0, 20))

        ctk.CTkLabel(linha2, text="Extensoes:", anchor="w",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left")
        self.entry_ext = ctk.CTkEntry(linha2, width=180)
        self.entry_ext.insert(0, ".pdf")
        self.entry_ext.pack(side="left", padx=(6, 0))

        # Linha 3: adicionar padrao (dropdown) e regex manual
        linha3 = ctk.CTkFrame(painel, fg_color="transparent")
        linha3.pack(fill="x", pady=(8, 2))
        ctk.CTkLabel(linha3, text="Padrao:", width=90, anchor="w",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left")
        self.dropdown = ctk.CTkOptionMenu(
            linha3, values=list(PADROES.keys()), width=230,
            fg_color=AZUL, button_color=AZUL_ESCURO, button_hover_color=AZUL_HOVER,
            dropdown_fg_color="#FFFFFF", dropdown_text_color=TEXTO,
            dropdown_hover_color=ZEBRA)
        self.dropdown.pack(side="left", padx=(0, 6))
        ctk.CTkButton(linha3, text="Adicionar", width=100, fg_color=AZUL,
                      hover_color=AZUL_HOVER,
                      command=self.adicionar_padrao).pack(side="left", padx=(0, 20))

        ctk.CTkLabel(linha3, text="Regex manual:", anchor="w",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left")
        self.entry_regex = ctk.CTkEntry(
            linha3, width=200, placeholder_text=r"ex.: \d{6}")
        self.entry_regex.pack(side="left", padx=(6, 6))
        self.entry_regex.bind("<Return>", lambda _e: self.adicionar_regex_manual())
        ctk.CTkButton(linha3, text="Adicionar", width=100, fg_color=AZUL,
                      hover_color=AZUL_HOVER,
                      command=self.adicionar_regex_manual).pack(side="left")

        # Linha 3b: padroes ativos (chips)
        linha3b = ctk.CTkFrame(painel, fg_color="transparent")
        linha3b.pack(fill="x", pady=(2, 2))
        ctk.CTkLabel(linha3b, text="Ativos:", width=90, anchor="nw",
                     text_color=TEXTO, font=("Segoe UI", 12)).pack(side="left", anchor="n")
        self.frame_chips = ctk.CTkScrollableFrame(
            linha3b, height=64, fg_color="#F7FAFC", corner_radius=6)
        self.frame_chips.pack(side="left", fill="x", expand=True)
        self.lbl_vazio = ctk.CTkLabel(
            self.frame_chips, text="Nenhum padrao ativo (remove apenas o prefixo).",
            text_color=CINZA, font=("Segoe UI", 11))
        self.lbl_vazio.pack(anchor="w", padx=6, pady=4)

        # Linha 4: opcao de prefixo
        linha4 = ctk.CTkFrame(painel, fg_color="transparent")
        linha4.pack(fill="x", pady=(2, 2))
        self.var_qualquer = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(linha4, text="Remover prefixo em qualquer posicao do nome",
                        variable=self.var_qualquer, text_color=TEXTO,
                        fg_color=AZUL, hover_color=AZUL_HOVER,
                        font=("Segoe UI", 11)).pack(side="left")

        # Linha 5: acoes
        linha5 = ctk.CTkFrame(painel, fg_color="transparent")
        linha5.pack(fill="x", pady=(4, 2))
        ctk.CTkButton(linha5, text="Aplicar renomeacao", width=170, fg_color=VINHO,
                      hover_color="#5C0030", command=self.aplicar).pack(side="right")
        ctk.CTkButton(linha5, text="Analisar", width=130, fg_color=AZUL,
                      hover_color=AZUL_HOVER, command=self.analisar).pack(side="right", padx=(0, 8))

    # ---------- Tabela ----------
    def _montar_tabela(self):
        wrapper = ctk.CTkFrame(self, fg_color="#FFFFFF")
        wrapper.pack(fill="both", expand=True, padx=16, pady=(6, 6))
        wrapper.grid_rowconfigure(0, weight=1)
        wrapper.grid_columnconfigure(0, weight=1)

        colunas = ("original", "novo", "status")
        self.tree = ttk.Treeview(wrapper, columns=colunas, show="headings", height=12)

        titulos = {"original": "Nome original", "novo": "Novo nome", "status": "Status"}
        larguras = {"original": 360, "novo": 360, "status": 130}
        for c in colunas:
            self.tree.heading(c, text=titulos[c], command=lambda col=c: self.ordenar(col))
            anchor = "w" if c != "status" else "center"
            self.tree.column(c, width=larguras[c], anchor=anchor, minwidth=90)

        self.titulos = titulos

        # Zebra + status
        self.tree.tag_configure("par", background="#FFFFFF")
        self.tree.tag_configure("impar", background=ZEBRA)
        self.tree.tag_configure("conflito", foreground=VINHO)
        self.tree.tag_configure("ignorado", foreground=CINZA)
        self.tree.tag_configure("feito", foreground="#00693E")

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<Double-1>", self.abrir_detalhe)

        scroll = AutoHideScrollbar(wrapper, width=11, command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)

    # ---------- Rodape ----------
    def _montar_rodape(self):
        rodape = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=34)
        rodape.pack(fill="x")
        rodape.pack_propagate(False)
        self.lbl_status = ctk.CTkLabel(
            rodape, text="Pronto.", text_color="#FFFFFF", font=("Segoe UI", 11),
        )
        self.lbl_status.pack(side="left", padx=16)
        ctk.CTkLabel(
            rodape, text="Igarape Digital", text_color="#E6F3FB",
            font=("Segoe UI", 11),
        ).pack(side="right", padx=16)

    # ==================== LOGICA ====================
    def escolher_pasta(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta")
        if caminho:
            self.entry_pasta.delete(0, "end")
            self.entry_pasta.insert(0, caminho)

    def _prefixos(self):
        return [p.strip() for p in self.entry_prefixos.get().split(",") if p.strip()]

    def _extensoes(self):
        return [e.strip().lower() for e in self.entry_ext.get().split(",") if e.strip()]

    # ---------- Gestao de padroes ----------
    def adicionar_padrao(self):
        nome = self.dropdown.get()
        self._add_pattern(nome, PADROES[nome])

    def adicionar_regex_manual(self):
        regex = self.entry_regex.get().strip()
        if not regex:
            return
        try:
            re.compile(regex)
        except re.error as e:
            messagebox.showerror("Regex invalido", f"Expressao invalida:\n\n{e}")
            return
        self._add_pattern(f"Manual: {regex}", regex)
        self.entry_regex.delete(0, "end")

    def _add_pattern(self, nome, regex):
        if any(p["regex"] == regex for p in self.patterns_ativos):
            return  # ja existe, evita duplicado
        self.patterns_ativos.append(
            {"nome": nome, "regex": regex, "compiled": re.compile(regex)})
        self._render_chips()

    def _remove_pattern(self, alvo):
        self.patterns_ativos = [p for p in self.patterns_ativos if p is not alvo]
        self._render_chips()

    def _render_chips(self):
        for w in self.frame_chips.winfo_children():
            w.destroy()

        if not self.patterns_ativos:
            self.lbl_vazio = ctk.CTkLabel(
                self.frame_chips,
                text="Nenhum padrao ativo (remove apenas o prefixo).",
                text_color=CINZA, font=("Segoe UI", 11))
            self.lbl_vazio.pack(anchor="w", padx=6, pady=4)
            return

        for p in self.patterns_ativos:
            chip = ctk.CTkFrame(self.frame_chips, fg_color="#FFFFFF", corner_radius=6)
            chip.pack(fill="x", padx=4, pady=2)
            ctk.CTkLabel(
                chip, text=f"{p['nome']}   ({p['regex']})",
                text_color=AZUL_ESCURO, font=("Segoe UI", 11), anchor="w",
            ).pack(side="left", padx=(8, 4), pady=2)
            ctk.CTkButton(
                chip, text="Remover", width=80, height=24,
                fg_color=VINHO, hover_color="#5C0030",
                font=("Segoe UI", 10),
                command=lambda alvo=p: self._remove_pattern(alvo),
            ).pack(side="right", padx=6, pady=2)

    def _novo_nome(self, nome):
        stem, ext = os.path.splitext(nome)  # nao mexe na extensao
        mudou = False

        # 1) Remove numeros de documento (CPF, PIS, regex personalizado)
        for pat in self.patterns:
            novo_stem = pat.sub("", stem)
            if novo_stem != stem:
                stem = novo_stem
                mudou = True

        # 2) Remove o prefixo literal (ex.: CPF_)
        qualquer = self.var_qualquer.get()
        for prefixo in self._prefixos():
            if qualquer and prefixo in stem:
                stem = stem.replace(prefixo, "", 1)
                mudou = True
                break
            if not qualquer and stem.startswith(prefixo):
                stem = stem[len(prefixo):]
                mudou = True
                break

        if not mudou:
            return nome

        # 3) Limpa separadores residuais (_, -, espaco) sobrando
        stem = re.sub(r"[ _]{2,}", "_", stem)
        stem = re.sub(r"-{2,}", "-", stem)
        stem = stem.strip(" _-.")

        # Seguranca: se a limpeza esvaziou o nome, mantem o original
        return (stem + ext) if stem else nome

    def analisar(self):
        pasta = Path(self.entry_pasta.get().strip())
        if not pasta.is_dir():
            messagebox.showerror("Erro", "Pasta nao encontrada. Selecione uma pasta valida.")
            return

        self.patterns = [p["compiled"] for p in self.patterns_ativos]

        exts = self._extensoes()
        arquivos = [
            p for p in pasta.iterdir()
            if p.is_file() and (not exts or p.suffix.lower() in exts)
        ]
        arquivos.sort(key=lambda p: p.name.lower())

        self.registros = []
        nomes_destino = set()

        for arq in arquivos:
            novo = self._novo_nome(arq.name)
            if novo == arq.name:
                status, motivo = "Ignorado", "Nenhum prefixo ou numero encontrado no nome."
            else:
                destino = arq.with_name(novo)
                if destino.exists() or novo in nomes_destino:
                    status, motivo = "Conflito", "Ja existe arquivo com o nome de destino."
                else:
                    status, motivo = "Renomear", "Pronto para renomear."
                    nomes_destino.add(novo)

            self.registros.append({
                "original": arq.name, "novo": novo,
                "status": status, "motivo": motivo, "path": arq,
            })

        self.coluna_ativa = None
        self.sort_ascending = {}
        self._resetar_titulos()
        self._render()
        self._atualizar_status()

    def aplicar(self):
        pendentes = [r for r in self.registros if r["status"] == "Renomear"]
        if not pendentes:
            messagebox.showinfo("Nada a fazer", "Nao ha arquivos para renomear. Rode 'Analisar' primeiro.")
            return

        confirma = messagebox.askyesno(
            "Confirmar",
            f"Renomear {len(pendentes)} arquivo(s)?\nEsta acao altera os arquivos na pasta.",
        )
        if not confirma:
            return

        renomeados, erros = 0, 0
        for r in self.registros:
            if r["status"] != "Renomear":
                continue
            destino = r["path"].with_name(r["novo"])
            try:
                if destino.exists():
                    r["status"], r["motivo"] = "Conflito", "Destino ja existia no momento da renomeacao."
                    erros += 1
                    continue
                os.rename(r["path"], destino)
                r["path"] = destino
                r["status"], r["motivo"] = "Renomeado", "Renomeado com sucesso."
                renomeados += 1
            except OSError as e:
                r["status"], r["motivo"] = "Conflito", f"Erro: {e}"
                erros += 1

        self._render()
        self.lbl_status.configure(
            text=f"Concluido: {renomeados} renomeado(s), {erros} com problema."
        )
        messagebox.showinfo("Concluido", f"{renomeados} arquivo(s) renomeado(s).")

    # ---------- Render / ordenacao ----------
    def _render(self):
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.registros):
            tags = ["par" if i % 2 == 0 else "impar"]
            if r["status"] == "Conflito":
                tags.append("conflito")
            elif r["status"] == "Ignorado":
                tags.append("ignorado")
            elif r["status"] == "Renomeado":
                tags.append("feito")
            self.tree.insert(
                "", "end", iid=str(i),
                values=(r["original"], r["novo"], r["status"]), tags=tags,
            )

    def _resetar_titulos(self):
        for c, t in self.titulos.items():
            self.tree.heading(c, text=t)

    def ordenar(self, coluna):
        # Primeiro clique em qualquer coluna = ascendente (padrao)
        if self.coluna_ativa != coluna:
            asc = True
        else:
            asc = not self.sort_ascending.get(coluna, True)

        self.coluna_ativa = coluna
        self.sort_ascending = {coluna: asc}

        self.registros.sort(key=lambda r: str(r[coluna]).lower(), reverse=not asc)

        self._resetar_titulos()
        seta = " \u2191" if asc else " \u2193"
        self.tree.heading(coluna, text=self.titulos[coluna] + seta)
        self._render()

    def abrir_detalhe(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        indice = int(sel[0])
        DetalheModal(self, self.registros[indice])

    def _atualizar_status(self):
        total = len(self.registros)
        renomear = sum(1 for r in self.registros if r["status"] == "Renomear")
        ignorados = sum(1 for r in self.registros if r["status"] == "Ignorado")
        conflitos = sum(1 for r in self.registros if r["status"] == "Conflito")
        self.lbl_status.configure(
            text=(f"Total: {total}  |  A renomear: {renomear}  |  "
                  f"Ignorados: {ignorados}  |  Conflitos: {conflitos}")
        )


if __name__ == "__main__":
    App().mainloop()

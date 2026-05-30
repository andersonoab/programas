import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import pyautogui
import pyperclip
import time
import re
import json
from pathlib import Path
from collections import defaultdict

try:
    from pynput import keyboard as pk
    HAS_PYNPUT = True
except Exception:
    HAS_PYNPUT = False

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

pyautogui.PAUSE = 0.0
pyautogui.FAILSAFE = True


class PreenchedorSSAF8Simples:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("CustomerThink | Preenchedor SSA F8 v4")
        self.app.geometry("1500x920")
        self.app.minsize(1180, 760)

        self.COR_HEADER = "#121C4E"
        self.COR_AZUL_1 = "#163A70"
        self.COR_AZUL_2 = "#245A9A"
        self.COR_AZUL_3 = "#3E7CC3"
        self.COR_AZUL_4 = "#DCE6F2"
        self.COR_AZUL_5 = "#EEF4FA"
        self.COR_SUCESSO = "#2F7D32"
        self.COR_ALERTA = "#C62828"
        self.COR_NEUTRA = "#6B7280"
        self.COR_INFO = "#406A9A"
        self.COR_CARD = "#FFFFFF"
        self.COR_FUNDO = "#F4F7FB"
        self.COR_BORDA = "#D7E0EA"
        self.COR_TEXTO = "#1F2937"
        self.COR_TEXTO_SUAVE = "#5F6B7A"
        self.COR_AMBAR = "#A97C00"

        # Estado SSA
        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.last_f8_time = 0.0
        self._debounce_secs = 0.30
        self._sending = False
        self.global_on = False
        self.hk_listener = None
        self._local_f8_bound = False

        # Estado Importação
        self.cadastro_df = None
        self.fatura_df = None
        self._fatura_caminho = ""
        self._resultados_cruzamento = []

        self.criar_interface()
        self._bind_local_keys()
        self.app.protocol("WM_DELETE_WINDOW", self.fechar_app)

        self.log("Aplicação inicializada — v4.")
        if HAS_PYNPUT:
            self.log("pynput disponível para hotkey global.")
        else:
            self.log("pynput não encontrado. F8 funcionará localmente na janela do app.")

    # =========================================================
    # BASE VISUAL
    # =========================================================
    def criar_card(self, parent):
        return ctk.CTkFrame(parent, fg_color=self.COR_CARD, corner_radius=14,
                            border_width=1, border_color=self.COR_BORDA)

    def criar_botao(self, parent, texto, comando, cor, hover, width=140):
        return ctk.CTkButton(parent, text=texto, command=comando,
                             fg_color=cor, hover_color=hover, text_color="white",
                             width=width, height=40, corner_radius=10,
                             font=("Segoe UI", 12, "bold"))

    def criar_entry(self, parent, width=None):
        kwargs = {"height": 40, "fg_color": "white", "border_color": "#AEBFD3",
                  "text_color": self.COR_TEXTO, "font": ("Segoe UI", 12)}
        if width is not None:
            kwargs["width"] = width
        return ctk.CTkEntry(parent, **kwargs)

    def criar_combo(self, parent, values, command=None, width=None):
        kwargs = {"values": values, "command": command, "state": "readonly",
                  "height": 40, "fg_color": "white", "border_color": "#AEBFD3",
                  "button_color": "#DCE6F2", "button_hover_color": "#C8D8EA",
                  "text_color": self.COR_TEXTO, "dropdown_fg_color": "white",
                  "dropdown_text_color": self.COR_TEXTO, "font": ("Segoe UI", 12)}
        if width is not None:
            kwargs["width"] = width
        return ctk.CTkComboBox(parent, **kwargs)

    # =========================================================
    # SCROLL CONTAINER
    # =========================================================
    def _frame_rolavel(self, parent):
        """Envolve o conteúdo da aba em Canvas com barras H+V.
        Expande para preencher quando a janela é maior;
        rola quando a janela é menor que o conteúdo."""
        parent.configure(fg_color=self.COR_FUNDO)
        vbar = ttk.Scrollbar(parent, orient="vertical")
        hbar = ttk.Scrollbar(parent, orient="horizontal")
        vbar.pack(side="right", fill="y")
        hbar.pack(side="bottom", fill="x")
        cv = tk.Canvas(parent, bg=self.COR_FUNDO, highlightthickness=0,
                       yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        cv.pack(side="left", fill="both", expand=True)
        vbar.configure(command=cv.yview)
        hbar.configure(command=cv.xview)
        inner = ctk.CTkFrame(cv, fg_color=self.COR_FUNDO, corner_radius=0)
        win_id = cv.create_window((0, 0), window=inner, anchor="nw")

        def _sync(e=None):
            cw, ch = cv.winfo_width(), cv.winfo_height()
            if cw < 2 or ch < 2:
                return
            rw, rh = inner.winfo_reqwidth(), inner.winfo_reqheight()
            w, h = max(rw, cw), max(rh, ch)
            cv.itemconfig(win_id, width=w, height=h)
            cv.configure(scrollregion=(0, 0, max(rw, cw), max(rh, ch)))

        inner.bind("<Configure>", lambda e: cv.after_idle(_sync))
        cv.bind("<Configure>", lambda e: cv.after_idle(_sync))
        # Scroll por botão do mouse (Linux: Button-4/5)
        cv.bind("<Button-4>", lambda e: cv.yview_scroll(-1, "units"))
        cv.bind("<Button-5>", lambda e: cv.yview_scroll(1, "units"))
        return inner

    # =========================================================
    # INTERFACE PRINCIPAL
    # =========================================================
    def criar_interface(self):
        self.app.configure(fg_color=self.COR_FUNDO)

        header = ctk.CTkFrame(self.app, fg_color=self.COR_HEADER, corner_radius=0, height=84)
        header.pack(fill="x")
        header.pack_propagate(False)

        esquerda = ctk.CTkFrame(header, fg_color="transparent")
        esquerda.pack(side="left", fill="both", expand=True, padx=20, pady=14)
        ctk.CTkLabel(esquerda, text="Preenchedor SSA | F8 + Importar Fatura",
                     font=("Segoe UI", 28, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(esquerda,
                     text="CustomerThink v4  |  Importa fatura Excel, cruza CPF × cadastro e preenche SSA via F8",
                     font=("Segoe UI", 13), text_color="#D9E7F5").pack(anchor="w", pady=(4, 0))

        direita = ctk.CTkFrame(header, fg_color="transparent")
        direita.pack(side="right", padx=20, pady=14)
        self.lbl_hotkey_mode = ctk.CTkLabel(direita, text="Hotkey: Local",
                                            font=("Segoe UI", 12, "bold"), text_color="white")
        self.lbl_hotkey_mode.pack(anchor="e")
        self.lbl_modo_execucao = ctk.CTkLabel(direita, text="Modo: F8 preenche e avança",
                                              font=("Segoe UI", 12), text_color="#D9E7F5")
        self.lbl_modo_execucao.pack(anchor="e", pady=(4, 0))

        self.tabs = ctk.CTkTabview(
            self.app, fg_color="transparent",
            segmented_button_fg_color=self.COR_AZUL_4,
            segmented_button_selected_color=self.COR_HEADER,
            segmented_button_selected_hover_color=self.COR_AZUL_1,
            segmented_button_unselected_color=self.COR_AZUL_4,
            segmented_button_unselected_hover_color=self.COR_AZUL_3,
            text_color="white", text_color_disabled=self.COR_TEXTO
        )
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        tab_imp = self.tabs.add("  Importar Fatura  ")
        tab_ssa = self.tabs.add("  Executar SSA  ")

        self.criar_aba_importar(tab_imp)
        self.criar_aba_ssa(tab_ssa)

        self.atualizar_resumo()
        self.atualizar_linha_atual()

    # =========================================================
    # ABA IMPORTAR FATURA
    # =========================================================
    def criar_aba_importar(self, parent):
        inner = self._frame_rolavel(parent)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_columnconfigure(1, weight=1)
        inner.grid_rowconfigure(0, weight=0, minsize=220)
        inner.grid_rowconfigure(1, weight=1)
        self._bloco_cadastro(inner)
        self._bloco_fatura(inner)
        self._bloco_cruzamento(inner)

    def _bloco_cadastro(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=0, padx=(0, 6), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(frame, text="1. Cadastro (SRA_Ativos)",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(frame, text="Planilha com CPF e CDC dos colaboradores ativos.",
                     font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE).pack(anchor="w", padx=16, pady=(0, 10))

        linha = ctk.CTkFrame(frame, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=(0, 6))
        self.criar_botao(linha, "Carregar cadastro", self.carregar_cadastro,
                         self.COR_HEADER, "#1B2C63", 165).pack(side="left", padx=(0, 12))
        self.lbl_cadastro_status = ctk.CTkLabel(linha, text="Nenhum arquivo carregado",
                                                font=("Segoe UI", 12), text_color=self.COR_NEUTRA)
        self.lbl_cadastro_status.pack(side="left")

        self.lbl_cadastro_info = ctk.CTkLabel(frame, text="", font=("Segoe UI", 11),
                                              text_color=self.COR_TEXTO_SUAVE)
        self.lbl_cadastro_info.pack(anchor="w", padx=16, pady=(0, 10))

        box = ctk.CTkFrame(frame, fg_color=self.COR_AZUL_5, corner_radius=10,
                           border_width=1, border_color=self.COR_BORDA)
        box.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkLabel(box, text="Mapeamento de colunas do cadastro",
                     font=("Segoe UI", 13, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=12, pady=(10, 6))

        g = ctk.CTkFrame(box, fg_color="transparent")
        g.pack(fill="x", padx=12, pady=(0, 12))
        g.grid_columnconfigure(0, weight=1)
        g.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(g, text="Coluna CPF", font=("Segoe UI", 12, "bold"),
                     text_color=self.COR_TEXTO).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
        ctk.CTkLabel(g, text="Coluna CDC", font=("Segoe UI", 12, "bold"),
                     text_color=self.COR_TEXTO).grid(row=0, column=1, sticky="w", pady=(0, 4))
        self.cmb_cad_cpf = self.criar_combo(g, [])
        self.cmb_cad_cpf.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.cmb_cad_cdc = self.criar_combo(g, [])
        self.cmb_cad_cdc.grid(row=1, column=1, sticky="ew")

    def _bloco_fatura(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=1, padx=(6, 0), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(frame, text="2. Fatura (qualquer fornecedor)",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 4))
        ctk.CTkLabel(frame, text="Selecione o Excel. Indique a coluna de CPF e a de Valor.",
                     font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE).pack(anchor="w", padx=16, pady=(0, 10))

        linha = ctk.CTkFrame(frame, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=(0, 6))
        self.criar_botao(linha, "Carregar fatura", self.carregar_fatura,
                         self.COR_AZUL_2, "#1E4F8A", 150).pack(side="left", padx=(0, 12))
        self.lbl_fatura_status = ctk.CTkLabel(linha, text="Nenhum arquivo carregado",
                                              font=("Segoe UI", 12), text_color=self.COR_NEUTRA)
        self.lbl_fatura_status.pack(side="left")

        self.lbl_fatura_info = ctk.CTkLabel(frame, text="", font=("Segoe UI", 11),
                                            text_color=self.COR_TEXTO_SUAVE)
        self.lbl_fatura_info.pack(anchor="w", padx=16, pady=(0, 10))

        box = ctk.CTkFrame(frame, fg_color=self.COR_AZUL_5, corner_radius=10,
                           border_width=1, border_color=self.COR_BORDA)
        box.pack(fill="x", padx=16, pady=(0, 16))
        ctk.CTkLabel(box, text="Mapeamento de colunas da fatura",
                     font=("Segoe UI", 13, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=12, pady=(10, 6))

        g = ctk.CTkFrame(box, fg_color="transparent")
        g.pack(fill="x", padx=12, pady=(0, 12))
        g.grid_columnconfigure(0, weight=1)
        g.grid_columnconfigure(1, weight=1)
        g.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(g, text="Aba (Excel) / ignorado em CSV", font=("Segoe UI", 12, "bold"),
                     text_color=self.COR_TEXTO).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 4))
        ctk.CTkLabel(g, text="Coluna CPF", font=("Segoe UI", 12, "bold"),
                     text_color=self.COR_TEXTO).grid(row=0, column=1, sticky="w", padx=(0, 8), pady=(0, 4))
        ctk.CTkLabel(g, text="Coluna Valor", font=("Segoe UI", 12, "bold"),
                     text_color=self.COR_TEXTO).grid(row=0, column=2, sticky="w", pady=(0, 4))
        self.cmb_fat_aba = self.criar_combo(g, [], command=self._ao_mudar_aba)
        self.cmb_fat_aba.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.cmb_fat_cpf = self.criar_combo(g, [])
        self.cmb_fat_cpf.grid(row=1, column=1, sticky="ew", padx=(0, 8))
        self.cmb_fat_valor = self.criar_combo(g, [])
        self.cmb_fat_valor.grid(row=1, column=2, sticky="ew")

    def _bloco_cruzamento(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=1, column=0, columnspan=2, padx=0, pady=0, sticky="nsew")
        frame.pack_propagate(True)

        # Cabeçalho
        topo = ctk.CTkFrame(frame, fg_color="transparent")
        topo.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(topo, text="3. Resultado do cruzamento",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(side="left")
        self.criar_botao(topo, "Cruzar e carregar para SSA", self.cruzar_e_carregar,
                         self.COR_SUCESSO, "#1B5E1F", 230).pack(side="right")
        self.criar_botao(topo, "Reenviar para SSA", self._reenviar_para_ssa,
                         self.COR_AZUL_2, "#1E4F8A", 165).pack(side="right", padx=(0, 8))

        self.lbl_cruzamento_resumo = ctk.CTkLabel(
            frame, text="Carregue o cadastro e a fatura, mapeie as colunas e clique em Cruzar.",
            font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE)
        self.lbl_cruzamento_resumo.pack(anchor="w", padx=16, pady=(0, 8))

        # ── KPI TOTAL ──────────────────────────────────────────
        self._kpi_frame = ctk.CTkFrame(frame, fg_color=self.COR_HEADER,
                                       corner_radius=12, height=88)
        self._kpi_frame.pack(fill="x", padx=16, pady=(0, 10))
        self._kpi_frame.pack_propagate(False)

        kpi_inner = ctk.CTkFrame(self._kpi_frame, fg_color="transparent")
        kpi_inner.place(relx=0.5, rely=0.5, anchor="center")

        self.lbl_kpi_titulo = ctk.CTkLabel(
            kpi_inner, text="TOTAL A LANÇAR NO SSA",
            font=("Segoe UI", 11, "bold"), text_color="#8BA8CC")
        self.lbl_kpi_titulo.pack()

        self.lbl_kpi_valor = ctk.CTkLabel(
            kpi_inner, text="—",
            font=("Segoe UI", 32, "bold"), text_color="white")
        self.lbl_kpi_valor.pack()

        self.lbl_kpi_detalhe = ctk.CTkLabel(
            kpi_inner, text="aguardando cruzamento",
            font=("Segoe UI", 11), text_color="#8BA8CC")
        self.lbl_kpi_detalhe.pack()
        # ───────────────────────────────────────────────────────

        # Grade
        tf = ctk.CTkFrame(frame, fg_color="transparent")
        tf.pack(fill="both", expand=True, padx=16, pady=(0, 4))

        cols = ("LINHA", "CPF", "NOME", "CDC", "VALOR", "STATUS")
        self.tree_cruzamento = ttk.Treeview(tf, columns=cols, show="headings")
        for col, txt, w, anc in [
            ("LINHA", "#", 44, "center"), ("CPF", "CPF", 130, "w"),
            ("NOME", "Nome", 210, "w"), ("CDC", "CDC", 100, "center"),
            ("VALOR", "Valor  (duplo clique p/ editar)", 185, "e"),
            ("STATUS", "Status", 90, "center")
        ]:
            self.tree_cruzamento.heading(col, text=txt)
            self.tree_cruzamento.column(col, width=w, anchor=anc, minwidth=40)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Cruz.Treeview", background="white", foreground="#1F2937",
                        rowheight=30, fieldbackground="white", font=("Segoe UI", 11),
                        borderwidth=0)
        style.configure("Cruz.Treeview.Heading", background="#E8EEF6",
                        foreground="#1F2937", font=("Segoe UI", 11, "bold"),
                        relief="flat", padding=(6, 6))
        style.map("Cruz.Treeview",
                  background=[("selected", "#C8DCF5")],
                  foreground=[("selected", "#121C4E")])
        self.tree_cruzamento.configure(style="Cruz.Treeview")

        self.tree_cruzamento.grid(row=0, column=0, sticky="nsew")
        sy = ttk.Scrollbar(tf, orient="vertical", command=self.tree_cruzamento.yview)
        sy.grid(row=0, column=1, sticky="ns")
        sx = ttk.Scrollbar(tf, orient="horizontal", command=self.tree_cruzamento.xview)
        sx.grid(row=1, column=0, sticky="ew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)
        self.tree_cruzamento.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

        self.tree_cruzamento.tag_configure("ok",      background="#F0FBF0", foreground="#1B5E1F")
        self.tree_cruzamento.tag_configure("ok_alt",  background="#E4F5E4", foreground="#1B5E1F")
        self.tree_cruzamento.tag_configure("ign",     background="#F9F9F9", foreground="#6B7280")
        self.tree_cruzamento.tag_configure("ign_alt", background="#F2F2F2", foreground="#6B7280")
        self.tree_cruzamento.bind("<Double-1>", self._editar_valor_inline)

        # Rodapé dica
        ctk.CTkLabel(frame, text="Duplo clique na coluna Valor para ajustar arredondamento ou taxa.",
                     font=("Segoe UI", 10), text_color=self.COR_TEXTO_SUAVE
                     ).pack(anchor="w", padx=16, pady=(2, 10))

    # =========================================================
    # LÓGICA DE IMPORTAÇÃO
    # =========================================================
    def carregar_cadastro(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Cadastro (SRA_Ativos)",
            filetypes=[
                ("Todos suportados", "*.xlsx *.xls *.xlsb *.csv"),
                ("Excel", "*.xlsx *.xls *.xlsb"),
                ("CSV", "*.csv"),
                ("Todos", "*.*")
            ],
            initialdir=str(Path.home() / "Documentos")
        )
        if not caminho:
            return
        try:
            if caminho.lower().endswith(".csv"):
                df = self._ler_csv_robusto(caminho)
            else:
                df = pd.read_excel(caminho, dtype=str)

            df.columns = [str(c).strip() for c in df.columns]
            self.cadastro_df = df
            cols = list(df.columns)
            self.cmb_cad_cpf.configure(values=cols)
            self.cmb_cad_cdc.configure(values=cols)
            for c in cols:
                cl = c.lower()
                if "cpf" in cl:
                    self.cmb_cad_cpf.set(c)
                if "cdc" in cl or ("centro" in cl and "custo" in cl):
                    self.cmb_cad_cdc.set(c)
            nome = Path(caminho).name
            self.lbl_cadastro_status.configure(text=nome, text_color=self.COR_SUCESSO)
            preview = ", ".join(cols[:6]) + ("..." if len(cols) > 6 else "")
            self.lbl_cadastro_info.configure(text=f"{len(df)} registros  |  colunas: {preview}")
            self.log(f"Cadastro: {nome} — {len(df)} registros.")
        except Exception as e:
            messagebox.showerror("Erro ao carregar cadastro", str(e))

    def carregar_fatura(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Fatura",
            filetypes=[
                ("Todos suportados", "*.xlsx *.xls *.xlsb *.csv"),
                ("Excel", "*.xlsx *.xls *.xlsb"),
                ("CSV", "*.csv"),
                ("Todos", "*.*")
            ]
        )
        if not caminho:
            return
        try:
            self._fatura_caminho = caminho
            if caminho.lower().endswith(".csv"):
                self.cmb_fat_aba.configure(values=["(CSV — sem abas)"])
                self.cmb_fat_aba.set("(CSV — sem abas)")
                self._carregar_csv(caminho)
            else:
                xl = pd.ExcelFile(caminho)
                abas = xl.sheet_names
                self.cmb_fat_aba.configure(values=abas)
                self.cmb_fat_aba.set(abas[0])
                self._carregar_aba(caminho, abas[0])
            self.lbl_fatura_status.configure(text=Path(caminho).name, text_color=self.COR_SUCESSO)
            self.log(f"Fatura: {Path(caminho).name}.")
        except Exception as e:
            messagebox.showerror("Erro ao carregar fatura", str(e))

    def _ler_csv_robusto(self, caminho):
        import io as _io
        raw = open(caminho, "rb").read()
        for enc in ["utf-8-sig", "latin-1", "cp1252", "utf-8"]:
            for sep in [";", ",", "\t"]:
                try:
                    df = pd.read_csv(_io.BytesIO(raw), sep=sep, dtype=str, encoding=enc)
                    if len(df.columns) > 1:
                        df.columns = [str(c).strip() for c in df.columns]
                        return df
                except Exception:
                    continue
        raise ValueError("Não foi possível ler o CSV. Verifique o arquivo.")

    def _carregar_csv(self, caminho):
        try:
            df = self._ler_csv_robusto(caminho)
            self._aplicar_fatura_df(df)
        except Exception as e:
            messagebox.showerror("Erro ao ler CSV", str(e))

    def _carregar_aba(self, caminho, aba):
        if aba == "(CSV — sem abas)":
            return
        try:
            df = pd.read_excel(caminho, sheet_name=aba, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            self._aplicar_fatura_df(df)
        except Exception as e:
            messagebox.showerror("Erro ao ler aba", str(e))

    def _aplicar_fatura_df(self, df):
        self.fatura_df = df
        cols = list(df.columns)
        self.cmb_fat_cpf.configure(values=cols)
        self.cmb_fat_valor.configure(values=cols)
        self.cmb_fat_cpf.set("")
        self.cmb_fat_valor.set("")
        for c in cols:
            cl = c.lower()
            if "cpf" in cl:
                self.cmb_fat_cpf.set(c)
            if any(k in cl for k in ["valor", "total", "fat", "vl_", "vr_"]):
                self.cmb_fat_valor.set(c)
        preview = ", ".join(cols[:6]) + ("..." if len(cols) > 6 else "")
        self.lbl_fatura_info.configure(text=f"{len(df)} linhas  |  colunas: {preview}")

    def _ao_mudar_aba(self, aba):
        if self._fatura_caminho and aba != "(CSV — sem abas)":
            self._carregar_aba(self._fatura_caminho, aba)

    def _norm_cpf(self, v):
        return re.sub(r"\D", "", str(v))

    def cruzar_e_carregar(self):
        if self.cadastro_df is None:
            messagebox.showwarning("Atenção", "Carregue o cadastro SRA_Ativos primeiro.")
            return
        if self.fatura_df is None:
            messagebox.showwarning("Atenção", "Carregue a fatura primeiro.")
            return

        col_cad_cpf = self.cmb_cad_cpf.get()
        col_cad_cdc = self.cmb_cad_cdc.get()
        col_fat_cpf = self.cmb_fat_cpf.get()
        col_fat_valor = self.cmb_fat_valor.get()

        if not col_cad_cpf or not col_cad_cdc:
            messagebox.showwarning("Atenção", "Mapeie as colunas do cadastro (CPF e CDC).")
            return
        if not col_fat_cpf or not col_fat_valor:
            messagebox.showwarning("Atenção", "Mapeie as colunas da fatura (CPF e Valor).")
            return

        cad = self.cadastro_df.copy()
        cad["_N"] = cad[col_cad_cpf].apply(self._norm_cpf)
        cpf_cdc = dict(zip(cad["_N"], cad[col_cad_cdc]))

        col_nome = next((c for c in self.fatura_df.columns if "nome" in c.lower()), None)

        ok_list, erro_list = [], []
        for _, row in self.fatura_df.iterrows():
            cpf_n = self._norm_cpf(row[col_fat_cpf])
            cpf_orig = str(row[col_fat_cpf]).strip()
            valor = str(row[col_fat_valor]).strip()
            nome = str(row[col_nome]).strip() if col_nome else ""
            if cpf_n and cpf_n in cpf_cdc:
                ok_list.append({"CPF": cpf_orig, "NOME": nome,
                                "CDC": str(cpf_cdc[cpf_n]).strip(), "VALOR": valor, "STATUS": "OK"})
            else:
                erro_list.append({"CPF": cpf_orig, "NOME": nome, "VALOR": valor})

        ignorar_todos = [False]
        erros_resolvidos = []
        for i, erro in enumerate(erro_list):
            if ignorar_todos[0]:
                erros_resolvidos.append({**erro, "CDC": "", "STATUS": "IGNORADO"})
                continue
            cdc_manual, acao = self._dialog_erro_cpf(
                erro["CPF"], erro["NOME"], erro["VALOR"], i + 1, len(erro_list))
            if acao == "ignorar_todos":
                ignorar_todos[0] = True
                erros_resolvidos.append({**erro, "CDC": "", "STATUS": "IGNORADO"})
            elif acao == "usar" and cdc_manual:
                erros_resolvidos.append({**erro, "CDC": cdc_manual, "STATUS": "OK"})
            else:
                erros_resolvidos.append({**erro, "CDC": "", "STATUS": "IGNORADO"})

        todos = ok_list + erros_resolvidos
        self._resultados_cruzamento = todos
        self._renderizar_cruzamento(todos)
        self._aglutinar_e_enviar_ssa(todos)

    def _aglutinar_e_enviar_ssa(self, todos):
        cdc_soma = defaultdict(float)
        cdc_qtd = defaultdict(int)
        for r in todos:
            if r["STATUS"] == "OK" and r["CDC"]:
                cdc = r["CDC"]
                cdc_qtd[cdc] += 1
                v = str(r["VALOR"]).strip().replace(".", "").replace(",", ".")
                try:
                    cdc_soma[cdc] += float(v)
                except ValueError:
                    pass
        registros = []
        for cdc, soma in cdc_soma.items():
            qtd = cdc_qtd[cdc]
            valor_fmt = f"{soma:.2f}".replace(".", ",")
            obs = f"Aglutinado: {qtd} lançamento(s)" if qtd > 1 else ""
            registros.append({"CDC": cdc, "DADO": valor_fmt, "TIPO": "VALOR",
                               "STATUS": "PENDENTE", "OBS": obs})
        n_ok = len(registros)
        n_ign = len([r for r in todos if r["STATUS"] == "IGNORADO"])

        if registros:
            self.df = pd.DataFrame(registros)
            self.pre_validar_dataframe()
            self.renderizar_base()
            self.atualizar_resumo()
            self.atualizar_linha_atual()

        msg = f"Cruzamento: {n_ok} carregado(s) para SSA | {n_ign} ignorado(s)."
        self.lbl_cruzamento_resumo.configure(
            text=msg, text_color=self.COR_SUCESSO if n_ok > 0 else self.COR_ALERTA)
        self.log(msg)

        if n_ok > 0:
            self.tabs.set("  Executar SSA  ")
            messagebox.showinfo("Pronto", f"{n_ok} linha(s) carregada(s) para o SSA.\n{n_ign} ignorada(s).")
        else:
            messagebox.showwarning("Atenção", "Nenhuma linha foi carregada. Verifique o mapeamento.")

    def _reenviar_para_ssa(self):
        if not self._resultados_cruzamento:
            messagebox.showwarning("Reenviar", "Realize o cruzamento primeiro.")
            return
        self._aglutinar_e_enviar_ssa(self._resultados_cruzamento)

    def _renderizar_cruzamento(self, resultados):
        for item in self.tree_cruzamento.get_children():
            self.tree_cruzamento.delete(item)
        ok_seq = 0
        ign_seq = 0
        for i, r in enumerate(resultados):
            if r["STATUS"] == "OK":
                tag = "ok" if ok_seq % 2 == 0 else "ok_alt"
                ok_seq += 1
            else:
                tag = "ign" if ign_seq % 2 == 0 else "ign_alt"
                ign_seq += 1
            self.tree_cruzamento.insert("", "end", tags=(tag,),
                values=(i + 1, r["CPF"], r["NOME"], r["CDC"], r["VALOR"], r["STATUS"]))
        self._atualizar_total()

    def _atualizar_total(self):
        total = 0.0
        n_ok = 0
        n_ign = 0
        for iid in self.tree_cruzamento.get_children():
            vals = self.tree_cruzamento.item(iid, "values")
            status = str(vals[5]) if len(vals) > 5 else ""
            if status == "OK":
                v = str(vals[4]).strip().replace(".", "").replace(",", ".")
                try:
                    total += float(v)
                    n_ok += 1
                except ValueError:
                    n_ok += 1
            else:
                n_ign += 1

        if n_ok > 0:
            total_fmt = f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            self.lbl_kpi_valor.configure(text=total_fmt, text_color="white")
            cor_kpi = "#0F5C2E" if n_ign == 0 else "#7A5500"
            self._kpi_frame.configure(fg_color=cor_kpi)
            self.lbl_kpi_detalhe.configure(
                text=f"{n_ok} registro(s) OK   ·   {n_ign} ignorado(s)   ·   {n_ok + n_ign} total",
                text_color="#A8D5B5" if n_ign == 0 else "#E8C97A")
        else:
            self._kpi_frame.configure(fg_color=self.COR_HEADER)
            self.lbl_kpi_valor.configure(text="—", text_color="white")
            self.lbl_kpi_detalhe.configure(text="aguardando cruzamento", text_color="#8BA8CC")

    def _editar_valor_inline(self, event):
        tree = self.tree_cruzamento
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col_id = tree.identify_column(event.x)
        iid = tree.identify_row(event.y)
        if not iid:
            return

        # Índice da coluna VALOR = #5 (1-based)
        cols = tree["columns"]
        col_idx = int(col_id.replace("#", "")) - 1
        if cols[col_idx] != "VALOR":
            return

        bbox = tree.bbox(iid, column=col_id)
        if not bbox:
            return

        x, y, w, h = bbox
        valor_atual = tree.item(iid, "values")[col_idx]

        entry = tk.Entry(tree, font=("Segoe UI", 11), justify="right",
                         relief="flat", bd=1, highlightthickness=1,
                         highlightbackground="#3E7CC3", highlightcolor="#3E7CC3")
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, valor_atual)
        entry.select_range(0, "end")
        entry.focus_set()

        def confirmar(e=None):
            novo = entry.get().strip()
            entry.destroy()
            if not novo:
                return
            vals = list(tree.item(iid, "values"))
            vals[col_idx] = novo
            tree.item(iid, values=vals)
            # Atualiza _resultados_cruzamento
            linha_idx = int(vals[0]) - 1
            if 0 <= linha_idx < len(self._resultados_cruzamento):
                self._resultados_cruzamento[linha_idx]["VALOR"] = novo
            self._atualizar_total()

        def cancelar(e=None):
            entry.destroy()

        entry.bind("<Return>", confirmar)
        entry.bind("<KP_Enter>", confirmar)
        entry.bind("<Escape>", cancelar)
        entry.bind("<FocusOut>", confirmar)

    def _dialog_erro_cpf(self, cpf, nome, valor, idx_erro, total_erros):
        dialog = ctk.CTkToplevel(self.app)
        dialog.title("CPF não encontrado")
        dialog.geometry("540x400")
        dialog.resizable(False, False)
        dialog.grab_set()
        dialog.focus_force()
        self.app.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 540) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 400) // 2
        dialog.geometry(f"540x400+{x}+{y}")

        resultado = {"cdc": "", "acao": "ignorar"}

        ctk.CTkLabel(dialog, text=f"Erro {idx_erro} de {total_erros}",
                     font=("Segoe UI", 12, "bold"), text_color=self.COR_ALERTA).pack(pady=(18, 2))
        ctk.CTkLabel(dialog, text="CPF não encontrado no cadastro",
                     font=("Segoe UI", 17, "bold"), text_color=self.COR_TEXTO).pack()

        info = ctk.CTkFrame(dialog, fg_color=self.COR_AZUL_5, corner_radius=10,
                            border_width=1, border_color=self.COR_BORDA)
        info.pack(fill="x", padx=24, pady=14)
        ctk.CTkLabel(info, text=f"CPF: {cpf}", font=("Segoe UI", 13),
                     text_color=self.COR_TEXTO).pack(anchor="w", padx=14, pady=(10, 2))
        if nome:
            ctk.CTkLabel(info, text=f"Nome: {nome}", font=("Segoe UI", 13),
                         text_color=self.COR_TEXTO).pack(anchor="w", padx=14, pady=2)
        ctk.CTkLabel(info, text=f"Valor: {valor}", font=("Segoe UI", 13),
                     text_color=self.COR_TEXTO).pack(anchor="w", padx=14, pady=(2, 10))

        ctk.CTkLabel(dialog,
                     text="Possível causa: colaborador desligado.\nInforme o CDC substituto ou ignore a linha:",
                     font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE, justify="center").pack(pady=(0, 8))

        entry_cdc = self.criar_entry(dialog)
        entry_cdc.pack(fill="x", padx=24, pady=(0, 14))
        entry_cdc.configure(placeholder_text="Digite o CDC para usar nesta linha...")

        bf = ctk.CTkFrame(dialog, fg_color="transparent")
        bf.pack(fill="x", padx=24, pady=(0, 18))

        def usar():
            cdc = entry_cdc.get().strip()
            if not cdc:
                messagebox.showwarning("Atenção", "Digite um CDC válido.", parent=dialog)
                return
            resultado["cdc"] = cdc
            resultado["acao"] = "usar"
            dialog.destroy()

        def ignorar():
            resultado["acao"] = "ignorar"
            dialog.destroy()

        def ignorar_todos():
            resultado["acao"] = "ignorar_todos"
            dialog.destroy()

        ctk.CTkButton(bf, text="Usar este CDC", command=usar,
                      fg_color=self.COR_SUCESSO, hover_color="#1B5E1F",
                      height=42, font=("Segoe UI", 13, "bold"), corner_radius=10
                      ).pack(side="left", expand=True, padx=(0, 4))
        ctk.CTkButton(bf, text="Ignorar esta linha", command=ignorar,
                      fg_color=self.COR_NEUTRA, hover_color="#5D6671",
                      height=42, font=("Segoe UI", 13), corner_radius=10
                      ).pack(side="left", expand=True, padx=4)
        if total_erros > 1:
            ctk.CTkButton(bf, text="Ignorar todos os erros", command=ignorar_todos,
                          fg_color=self.COR_ALERTA, hover_color="#AE2323",
                          height=42, font=("Segoe UI", 13), corner_radius=10
                          ).pack(side="left", expand=True, padx=(4, 0))

        dialog.wait_window()
        return resultado["cdc"], resultado["acao"]

    # =========================================================
    # ABA EXECUTAR SSA
    # =========================================================
    def criar_aba_ssa(self, parent):
        inner = self._frame_rolavel(parent)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_columnconfigure(1, weight=1)
        inner.grid_rowconfigure(0, weight=0, minsize=260)
        inner.grid_rowconfigure(1, weight=1)
        self.criar_bloco_entrada(inner)
        self.criar_bloco_config(inner)
        self.criar_bloco_preview(inner)
        self.criar_bloco_execucao(inner)

    # =========================================================
    # BLOCO ENTRADA
    # =========================================================
    def criar_bloco_entrada(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=0, padx=(0, 6), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(frame, text="1. Entrada rápida",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 8))

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(fill="x", padx=16, pady=(0, 8))
        self.criar_botao(botoes, "Ler dados colados", self.ler_dados_colados, self.COR_HEADER, "#1B2C63", 165).pack(side="left", padx=(0, 8))
        self.criar_botao(botoes, "Modelo valor", self.inserir_modelo_valor, self.COR_AZUL_2, "#1E4F8A", 130).pack(side="left", padx=(0, 8))
        self.criar_botao(botoes, "Modelo percentual", self.inserir_modelo_percentual, self.COR_AZUL_3, "#2D68A8", 160).pack(side="left", padx=(0, 8))
        self.criar_botao(botoes, "Limpar texto", self.limpar_texto, self.COR_NEUTRA, "#5D6671", 120).pack(side="left", padx=(0, 8))
        self.criar_botao(botoes, "Limpar base", self.limpar_base, "#A83E4A", "#933540", 120).pack(side="left", padx=(0, 8))
        self.criar_botao(botoes, "Validar base", self.validar_base_visual, self.COR_INFO, "#355D87", 125).pack(side="left")

        ctk.CTkLabel(frame, text="Aceita TAB, ponto e vírgula ou múltiplos espaços. Pode colar com ou sem cabeçalho.",
                     font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE).pack(anchor="w", padx=16, pady=(0, 6))

        tf = ctk.CTkFrame(frame, fg_color="transparent")
        tf.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.txt_entrada = ctk.CTkTextbox(tf, height=260, font=("Consolas", 12),
                                          fg_color="#FBFCFE", text_color=self.COR_TEXTO,
                                          border_width=1, border_color=self.COR_BORDA)
        self.txt_entrada.pack(side="left", fill="both", expand=True)
        self.scroll_txt = ctk.CTkScrollbar(tf, command=self.txt_entrada.yview)
        self.scroll_txt.pack(side="right", fill="y")
        self.txt_entrada.configure(yscrollcommand=self.scroll_txt.set)

    # =========================================================
    # BLOCO CONFIG
    # =========================================================
    def criar_bloco_config(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=0, column=1, padx=(6, 0), pady=(0, 8), sticky="nsew")

        ctk.CTkLabel(frame, text="2. Configuração operacional",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 8))

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=16, pady=(0, 10))
        for i in range(5):
            grid.grid_columnconfigure(i, weight=1)

        for i, txt in enumerate(["Usar campo", "Navegação", "Delay entre ações", "Delay inicial", "Qtd. a executar"]):
            ctk.CTkLabel(grid, text=txt, font=("Segoe UI", 12, "bold"),
                         text_color=self.COR_TEXTO).grid(row=0, column=i, sticky="w", padx=6, pady=(0, 4))

        self.cmb_usar = self.criar_combo(grid, ["VALOR", "PERCENTUAL"])
        self.cmb_usar.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_usar.set("VALOR")

        self.cmb_modo = self.criar_combo(grid, ["TAB"])
        self.cmb_modo.grid(row=1, column=1, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_modo.set("TAB")

        self.entry_delay = self.criar_entry(grid)
        self.entry_delay.grid(row=1, column=2, sticky="ew", padx=6, pady=(0, 8))
        self.entry_delay.insert(0, "0,15")

        self.entry_delay_inicial = self.criar_entry(grid)
        self.entry_delay_inicial.grid(row=1, column=3, sticky="ew", padx=6, pady=(0, 8))
        self.entry_delay_inicial.insert(0, "3")

        self.entry_qtd = self.criar_entry(grid)
        self.entry_qtd.grid(row=1, column=4, sticky="ew", padx=6, pady=(0, 8))
        self.entry_qtd.insert(0, "TODOS")

        for i, txt in enumerate(["TAB após CDC", "TAB após dado", "Limpar campo atual", "Ação final", "Linha inicial"]):
            ctk.CTkLabel(grid, text=txt, font=("Segoe UI", 12, "bold"),
                         text_color=self.COR_TEXTO).grid(row=2, column=i, sticky="w", padx=6, pady=(8, 4))

        self.entry_tabs_cdc = self.criar_entry(grid)
        self.entry_tabs_cdc.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 8))
        self.entry_tabs_cdc.insert(0, "3")

        self.entry_tabs_final = self.criar_entry(grid)
        self.entry_tabs_final.grid(row=3, column=1, sticky="ew", padx=6, pady=(0, 8))
        self.entry_tabs_final.insert(0, "0")

        self.cmb_limpar = self.criar_combo(grid, ["SIM", "NAO"])
        self.cmb_limpar.grid(row=3, column=2, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_limpar.set("SIM")

        self.cmb_acao_final = self.criar_combo(grid, ["ENTER", "TAB", "NENHUMA"])
        self.cmb_acao_final.grid(row=3, column=3, sticky="ew", padx=6, pady=(0, 8))
        self.cmb_acao_final.set("ENTER")

        self.entry_linha_inicial = self.criar_entry(grid)
        self.entry_linha_inicial.grid(row=3, column=4, sticky="ew", padx=6, pady=(0, 8))
        self.entry_linha_inicial.insert(0, "1")

        linha2 = ctk.CTkFrame(frame, fg_color="transparent")
        linha2.pack(fill="x", padx=16, pady=(0, 10))
        self.criar_botao(linha2, "Salvar layout", self.salvar_layout_json, self.COR_AZUL_2, "#1E4F8A", 130).pack(side="left", padx=(0, 8))
        self.criar_botao(linha2, "Carregar layout", self.carregar_layout_json, self.COR_AZUL_3, "#2D68A8", 145).pack(side="left", padx=(0, 8))
        self.criar_botao(linha2, "Exportar log", self.exportar_log_txt, self.COR_INFO, "#355D87", 130).pack(side="left", padx=(0, 8))
        self.criar_botao(linha2, "Ignorar selecionada", self.marcar_linha_ignorada, self.COR_NEUTRA, "#59616D", 165).pack(side="left")

        box = ctk.CTkFrame(frame, fg_color=self.COR_AZUL_5, corner_radius=10, border_width=1, border_color=self.COR_BORDA)
        box.pack(fill="x", padx=16, pady=(0, 10))
        ctk.CTkLabel(box, text="Lógica simplificada para 1 tela",
                     font=("Segoe UI", 14, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=12, pady=(12, 6))
        ctk.CTkLabel(box,
                     text=("1. O usuário deixa o cursor no primeiro campo do SSA.\n"
                            "2. Pressiona F8.\n"
                            "3. O app escreve CDC, navega por TAB e escreve o valor ou percentual.\n"
                            "4. Depois marca a linha como OK e avança automaticamente.\n"
                            "5. Não depende de capturar posição na tela."),
                     justify="left", font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE
                     ).pack(anchor="w", padx=12, pady=(0, 12))

    # =========================================================
    # BLOCO PREVIEW
    # =========================================================
    def criar_bloco_preview(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=1, column=0, padx=(0, 6), pady=(0, 0), sticky="nsew")

        ctk.CTkLabel(frame, text="3. Preview da base",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 8))

        self.lbl_resumo = ctk.CTkLabel(frame, text="Total: 0 | Pendentes: 0 | OK: 0 | Erro: 0 | Ignorados: 0",
                                       font=("Segoe UI", 12, "bold"), text_color=self.COR_TEXTO_SUAVE)
        self.lbl_resumo.pack(anchor="w", padx=16, pady=(0, 8))

        tf = ctk.CTkFrame(frame, fg_color="transparent")
        tf.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.tree = ttk.Treeview(tf, columns=("LINHA", "CDC", "DADO", "TIPO", "STATUS", "OBS"), show="headings")
        for col, txt, w, anc in [
            ("LINHA", "Linha", 70, "center"), ("CDC", "CDC", 120, "center"),
            ("DADO", "Dado", 140, "e"), ("TIPO", "Tipo", 110, "center"),
            ("STATUS", "Status", 120, "center"), ("OBS", "Observação", 330, "w")
        ]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor=anc)

        self.tree.grid(row=0, column=0, sticky="nsew")
        sy = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        sy.grid(row=0, column=1, sticky="ns")
        sx = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        sx.grid(row=1, column=0, sticky="ew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)

    # =========================================================
    # BLOCO EXECUÇÃO
    # =========================================================
    def criar_bloco_execucao(self, parent):
        frame = self.criar_card(parent)
        frame.grid(row=1, column=1, padx=(6, 0), pady=(0, 0), sticky="nsew")

        ctk.CTkLabel(frame, text="4. Execução por hotkey",
                     font=("Segoe UI", 18, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(14, 8))

        ctk.CTkLabel(frame,
                     text=("Fluxo operacional:\n"
                            "1. Cole os dados.\n"
                            "2. Leia e valide a base.\n"
                            "3. Ajuste os TABs conforme o SSA.\n"
                            "4. Clique no primeiro campo do SSA.\n"
                            "5. Pressione F8 para preencher a linha atual.\n"
                            "6. Use F6 para pular e F9 para voltar."),
                     justify="left", font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE
                     ).pack(anchor="w", padx=16, pady=(0, 10))

        b1 = ctk.CTkFrame(frame, fg_color="transparent")
        b1.pack(fill="x", padx=16, pady=(0, 8))
        self.criar_botao(b1, "Iniciar da linha", self.iniciar_da_linha_configurada, self.COR_AZUL_2, "#1E4F8A", 135).pack(side="left", padx=(0, 8))
        self.criar_botao(b1, "Executar 1 linha", self.executar_proxima_linha_manual, self.COR_AZUL_3, "#2D68A8", 135).pack(side="left", padx=(0, 8))
        self.criar_botao(b1, "Pular linha", self.pular_linha_manual, self.COR_AMBAR, "#916B00", 120).pack(side="left", padx=(0, 8))
        self.criar_botao(b1, "Voltar linha", self.voltar_linha, self.COR_NEUTRA, "#5D6671", 125).pack(side="left", padx=(0, 8))

        b2 = ctk.CTkFrame(frame, fg_color="transparent")
        b2.pack(fill="x", padx=16, pady=(0, 10))
        self.btn_hotkey = self.criar_botao(b2, "Ativar hotkey global", self.toggle_global, self.COR_HEADER, "#1B2C63", 180)
        self.btn_hotkey.pack(side="left", padx=(0, 8))
        self.criar_botao(b2, "Marcar erro", self.marcar_erro_manual, self.COR_ALERTA, "#AE2323", 120).pack(side="left", padx=(0, 8))

        self.lbl_status = ctk.CTkLabel(frame, text="Status: aguardando dados.",
                                       font=("Segoe UI", 12, "bold"), text_color=self.COR_HEADER)
        self.lbl_status.pack(anchor="w", padx=16, pady=(0, 8))
        self.lbl_linha = ctk.CTkLabel(frame, text="Linha atual: 0",
                                      font=("Segoe UI", 12), text_color=self.COR_TEXTO_SUAVE)
        self.lbl_linha.pack(anchor="w", padx=16, pady=(0, 4))
        self.lbl_progresso = ctk.CTkLabel(frame, text="Progresso: 0/0",
                                          font=("Segoe UI", 12, "bold"), text_color=self.COR_INFO)
        self.lbl_progresso.pack(anchor="w", padx=16, pady=(0, 8))

        ctk.CTkLabel(frame, text="Log operacional",
                     font=("Segoe UI", 13, "bold"), text_color=self.COR_TEXTO).pack(anchor="w", padx=16, pady=(0, 6))

        lf = ctk.CTkFrame(frame, fg_color="transparent")
        lf.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.txt_log = ctk.CTkTextbox(lf, height=440, font=("Consolas", 11),
                                      fg_color="#FBFCFE", text_color=self.COR_TEXTO,
                                      border_width=1, border_color=self.COR_BORDA)
        self.txt_log.pack(side="left", fill="both", expand=True)
        self.scroll_log = ctk.CTkScrollbar(lf, command=self.txt_log.yview)
        self.scroll_log.pack(side="right", fill="y")
        self.txt_log.configure(yscrollcommand=self.scroll_log.set)

    # =========================================================
    # LOG E STATUS
    # =========================================================
    def log(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        linha = f"[{timestamp}] {msg}\n"
        if hasattr(self, "txt_log"):
            self.txt_log.insert("end", linha)
            self.txt_log.see("end")
            self.app.update_idletasks()
        else:
            print(linha)

    def set_status(self, msg, cor=None):
        self.lbl_status.configure(text=f"Status: {msg}", text_color=cor or self.COR_HEADER)
        self.app.update_idletasks()

    # =========================================================
    # HOTKEYS
    # =========================================================
    def _bind_local_keys(self):
        self.app.bind("<F8>", self._on_f8_local)
        self.app.bind("<F6>", lambda e: self.on_pular_hotkey())
        self.app.bind("<F9>", lambda e: self.on_voltar_hotkey())
        self._local_f8_bound = True

    def _unbind_local_f8(self):
        if self._local_f8_bound:
            try:
                self.app.unbind("<F8>")
            except Exception:
                pass
            self._local_f8_bound = False

    def _has_modifiers(self, event):
        st = getattr(event, "state", 0)
        return bool(st & 0x0001 or st & 0x0004 or st & 0x0008)

    def _on_f8_local(self, event):
        if self._has_modifiers(event):
            return
        self.executar_proxima_linha_manual()

    def toggle_global(self):
        if not HAS_PYNPUT:
            messagebox.showerror("Hotkey global", "Instale o pacote 'pynput' para usar F8 global.")
            return
        if not self.global_on:
            try:
                self._unbind_local_f8()
                mapping = {
                    "<f8>": lambda: self.app.after(0, self.executar_proxima_linha_manual),
                    "<f6>": lambda: self.app.after(0, self.on_pular_hotkey),
                    "<f9>": lambda: self.app.after(0, self.on_voltar_hotkey),
                }
                self.hk_listener = pk.GlobalHotKeys(mapping)
                self.hk_listener.start()
                self.global_on = True
                self.btn_hotkey.configure(text="Desativar hotkey global", fg_color="#6c757d", hover_color="#5c636a")
                self.lbl_hotkey_mode.configure(text="Hotkey: Global")
                self.log("Hotkey global ativada: F8 preenche, F6 pula, F9 volta.")
                self.set_status("hotkey global ativada", self.COR_SUCESSO)
            except Exception as e:
                messagebox.showerror("Hotkey global", f"Falha ao ativar.\n\n{e}")
        else:
            try:
                if self.hk_listener:
                    self.hk_listener.stop()
            except Exception:
                pass
            self.hk_listener = None
            self.global_on = False
            self.app.bind("<F8>", self._on_f8_local)
            self._local_f8_bound = True
            self.btn_hotkey.configure(text="Ativar hotkey global", fg_color=self.COR_HEADER, hover_color="#1B2C63")
            self.lbl_hotkey_mode.configure(text="Hotkey: Local")
            self.log("Hotkey global desativada.")
            self.set_status("hotkey global desativada", self.COR_INFO)

    # =========================================================
    # ENTRADA / PARSER
    # =========================================================
    def inserir_modelo_valor(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert("1.0", "CDC\tVALOR\n300000\t843,60\n301000\t3532,00\n303000\t2000,00")
        self.cmb_usar.set("VALOR")

    def inserir_modelo_percentual(self):
        self.txt_entrada.delete("1.0", "end")
        self.txt_entrada.insert("1.0", "CDC\tPERCENTUAL\n300000\t25\n301000\t35\n303000\t40")
        self.cmb_usar.set("PERCENTUAL")

    def limpar_texto(self):
        self.txt_entrada.delete("1.0", "end")

    def limpar_base(self):
        self.df = pd.DataFrame(columns=["CDC", "DADO", "TIPO", "STATUS", "OBS"])
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Base limpa.")
        self.set_status("base limpa", self.COR_INFO)

    def detectar_cabecalho(self, primeira_linha):
        p1 = str(primeira_linha[0]).strip().lower()
        p2 = str(primeira_linha[1]).strip().lower()
        palavras = ["cdc", "valor", "percentual", "percent", "%", "dado"]
        return any(p in p1 for p in palavras) or any(p in p2 for p in palavras)

    def separar_linha(self, linha):
        linha = linha.strip()
        if not linha:
            return None
        if "\t" in linha:
            partes = [p.strip() for p in linha.split("\t") if p.strip()]
            if len(partes) >= 2:
                return partes[0], partes[1]
        if ";" in linha:
            partes = [p.strip() for p in linha.split(";") if p.strip()]
            if len(partes) >= 2:
                return partes[0], partes[1]
        partes = [p.strip() for p in re.split(r"\s{2,}", linha) if p.strip()]
        if len(partes) >= 2:
            return partes[0], partes[1]
        return None

    def ler_dados_colados(self):
        texto = self.txt_entrada.get("1.0", "end").strip()
        if not texto:
            messagebox.showwarning("Atenção", "Cole os dados antes de ler.")
            return
        linhas = [l for l in texto.splitlines() if l.strip()]
        if not linhas:
            messagebox.showwarning("Atenção", "Não há linhas válidas.")
            return
        primeira = self.separar_linha(linhas[0])
        inicio = 1 if primeira and self.detectar_cabecalho(primeira) else 0
        tipo = self.cmb_usar.get()
        registros = []
        for i in range(inicio, len(linhas)):
            r = self.separar_linha(linhas[i])
            if not r:
                self.log(f"Linha ignorada: {linhas[i]}")
                continue
            cdc, dado = r
            registros.append({"CDC": cdc, "DADO": dado, "TIPO": tipo, "STATUS": "PENDENTE", "OBS": ""})
        if not registros:
            messagebox.showwarning("Atenção", "Nenhum dado válido identificado.")
            return
        self.df = pd.DataFrame(registros)
        self.pre_validar_dataframe()
        self.iniciar_da_linha_configurada(silencioso=True)
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log(f"{len(self.df)} linha(s) carregada(s) por cola.")
        self.set_status(f"{len(self.df)} linha(s) carregada(s)", self.COR_SUCESSO)

    # =========================================================
    # PRÉ-VALIDAÇÃO
    # =========================================================
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
                self.df.at[idx, "OBS"] = ""

    def validar_base_visual(self):
        if self.df.empty:
            messagebox.showwarning("Validação", "Não há base carregada.")
            return
        self.pre_validar_dataframe()
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log("Base validada.")
        self.set_status("base validada", self.COR_INFO)

    # =========================================================
    # RESUMO E GRID
    # =========================================================
    def renderizar_base(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        if self.df.empty:
            return
        for idx, row in self.df.iterrows():
            self.tree.insert("", "end", iid=str(idx),
                             values=(idx + 1, row["CDC"], row["DADO"], row["TIPO"], row["STATUS"], row["OBS"]))

    def atualizar_status_linha(self, idx, status, obs=None):
        if idx < 0 or idx >= len(self.df):
            return
        self.df.at[idx, "STATUS"] = status
        if obs is not None:
            self.df.at[idx, "OBS"] = obs
        self.tree.item(str(idx), values=(
            idx + 1, self.df.at[idx, "CDC"], self.df.at[idx, "DADO"],
            self.df.at[idx, "TIPO"], self.df.at[idx, "STATUS"], self.df.at[idx, "OBS"]))
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
        self.lbl_resumo.configure(text=f"Total: {total} | Pendentes: {pend} | OK: {ok} | Erro: {erro} | Ignorados: {ign}")
        self.lbl_progresso.configure(text=f"Progresso: {ok + ign}/{total}")

    def obter_proxima_linha_pendente(self):
        if self.df.empty:
            return None
        pendentes = self.df.index[self.df["STATUS"].isin(["PENDENTE", "ERRO"])]
        return int(pendentes[0]) if len(pendentes) > 0 else None

    def atualizar_linha_atual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.lbl_linha.configure(text="Linha atual: concluído")
        else:
            self.lbl_linha.configure(text=f"Linha atual: {idx + 1} de {len(self.df)}")
            try:
                self.tree.selection_set(str(idx))
                self.tree.focus(str(idx))
                self.tree.see(str(idx))
            except Exception:
                pass

    # =========================================================
    # PARÂMETROS
    # =========================================================
    def obter_delay(self):
        try:
            return float(str(self.entry_delay.get()).replace(",", "."))
        except Exception:
            return 0.10

    def obter_delay_inicial(self):
        try:
            return int(float(str(self.entry_delay_inicial.get()).replace(",", ".")))
        except Exception:
            return 2

    def obter_qtd(self):
        txt = str(self.entry_qtd.get()).strip().upper()
        if txt in ("", "TODOS"):
            return len(self.df)
        try:
            return max(1, min(int(txt), len(self.df)))
        except Exception:
            return len(self.df)

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
    # EXECUÇÃO
    # =========================================================
    def validar_execucao(self):
        if self.df.empty:
            messagebox.showwarning("Validação", "Leia os dados antes de executar.")
            return False
        return True

    def colar_texto(self, texto, limpar=False):
        pyperclip.copy(str(texto))
        if limpar:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.08)
        pyautogui.hotkey("ctrl", "v")

    def digitar_texto(self, texto, intervalo=0.02, limpar=False):
        if limpar:
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.08)
        pyautogui.write(str(texto), interval=intervalo)

    def limpar_campo_atual(self, delay=0.08):
        pyautogui.hotkey("ctrl", "a")
        time.sleep(delay)
        pyautogui.press("backspace")
        time.sleep(delay)

    def preparar_dado_para_digitacao(self, dado, usar):
        valor = str(dado).strip()
        if usar == "VALOR":
            valor = valor.replace(" ", "")
        elif usar == "PERCENTUAL":
            valor = valor.replace("%", "").strip()
        return valor

    def escrever_cdc(self, cdc):
        self.colar_texto(cdc, limpar=False)

    def escrever_dado_no_campo(self, dado, usar, delay):
        valor = self.preparar_dado_para_digitacao(dado, usar)
        if self.cmb_limpar.get() == "SIM":
            self.limpar_campo_atual(delay=max(0.05, delay))
        self.digitar_texto(valor, intervalo=max(0.01, delay / 5), limpar=False)
        return valor

    def executar_tabs(self, quantidade, delay):
        for _ in range(quantidade):
            pyautogui.press("tab")
            time.sleep(delay)

    def executar_acao_final(self, delay):
        acao = self.cmb_acao_final.get()
        if acao == "ENTER":
            pyautogui.press("enter")
            time.sleep(delay)
        elif acao == "TAB":
            pyautogui.press("tab")
            time.sleep(delay)

    def iniciar_da_linha_configurada(self, silencioso=False):
        if self.df.empty:
            if not silencioso:
                messagebox.showwarning("Início", "Carregue a base primeiro.")
            return
        linha = self.obter_linha_inicial() - 1
        for idx in range(len(self.df)):
            if idx < linha and self.df.at[idx, "STATUS"] in ["PENDENTE", "ERRO"]:
                self.df.at[idx, "STATUS"] = "IGNORADO"
                self.df.at[idx, "OBS"] = "Ignorado por linha inicial"
        for idx in range(linha, len(self.df)):
            if self.df.at[idx, "STATUS"] not in ["OK", "IGNORADO"]:
                self.df.at[idx, "STATUS"] = "PENDENTE"
                if self.df.at[idx, "OBS"] == "Ignorado por linha inicial":
                    self.df.at[idx, "OBS"] = ""
        self.renderizar_base()
        self.atualizar_resumo()
        self.atualizar_linha_atual()
        self.log(f"Posicionado na linha {linha + 1}.")
        self.set_status(f"linha inicial: {linha + 1}", self.COR_INFO)

    def executar_linha(self, idx):
        if idx is None or idx >= len(self.df) or self._sending:
            return False
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
            tabs_cdc = self.obter_tabs_cdc()
            tabs_final = self.obter_tabs_final()

            if not cdc:
                self.atualizar_status_linha(idx, "IGNORADO", "CDC vazio")
                self.atualizar_linha_atual()
                return True
            if not dado:
                self.atualizar_status_linha(idx, "IGNORADO", "Dado vazio")
                self.atualizar_linha_atual()
                return True

            delay_inicial = self.obter_delay_inicial()
            if delay_inicial > 0:
                self.set_status(f"aguardando {delay_inicial}s...", self.COR_INFO)
                time.sleep(delay_inicial)

            self.escrever_cdc(cdc)
            time.sleep(delay)
            self.executar_tabs(tabs_cdc, delay)
            self.escrever_dado_no_campo(dado, usar, delay)
            time.sleep(delay)
            self.executar_tabs(tabs_final, delay)
            self.executar_acao_final(delay)

            self.atualizar_status_linha(idx, "OK", "Preenchido com F8")
            self.log(f"Linha {idx + 1}: OK.")
            self.set_status(f"linha {idx + 1} preenchida", self.COR_SUCESSO)
            self.atualizar_linha_atual()
            return True
        except Exception as e:
            self.atualizar_status_linha(idx, "ERRO", str(e))
            self.log(f"Linha {idx + 1}: erro — {e}")
            self.set_status(f"erro na linha {idx + 1}", self.COR_ALERTA)
            self.atualizar_linha_atual()
            return False
        finally:
            time.sleep(0.02)
            self._sending = False

    def executar_proxima_linha_manual(self):
        if not self.validar_execucao():
            return
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.set_status("não há linhas pendentes", self.COR_INFO)
            return
        self.executar_linha(idx)

    def pular_linha_manual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.set_status("não há linhas pendentes", self.COR_INFO)
            return
        self.atualizar_status_linha(idx, "IGNORADO", "Pulado manualmente")
        self.log(f"Linha {idx + 1} pulada.")
        self.set_status(f"linha {idx + 1} pulada", self.COR_AMBAR)
        self.atualizar_linha_atual()

    def marcar_erro_manual(self):
        idx = self.obter_proxima_linha_pendente()
        if idx is None:
            self.set_status("não há linhas pendentes", self.COR_INFO)
            return
        self.atualizar_status_linha(idx, "ERRO", "Marcado manualmente")
        self.log(f"Linha {idx + 1} marcada como erro.")
        self.set_status(f"linha {idx + 1} com erro", self.COR_ALERTA)
        self.atualizar_linha_atual()

    def on_pular_hotkey(self):
        self.pular_linha_manual()

    def on_voltar_hotkey(self):
        self.voltar_linha()

    def voltar_linha(self):
        if self.df.empty:
            return
        linhas = self.df.index[self.df["STATUS"].isin(["OK", "IGNORADO", "ERRO"])].tolist()
        if not linhas:
            self.set_status("nenhuma linha concluída para retornar", self.COR_INFO)
            return
        idx = linhas[-1]
        self.atualizar_status_linha(idx, "PENDENTE", "Retornada manualmente")
        self.log(f"Linha {idx + 1} retornada para PENDENTE.")
        self.set_status(f"linha {idx + 1} voltou para pendente", self.COR_INFO)
        self.atualizar_linha_atual()

    def marcar_linha_ignorada(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Ignorar", "Selecione uma linha na grade.")
            return
        idx = int(sel[0])
        self.atualizar_status_linha(idx, "IGNORADO", "Ignorada manualmente")
        self.log(f"Linha {idx + 1} ignorada.")
        self.set_status(f"linha {idx + 1} ignorada", self.COR_INFO)
        self.atualizar_linha_atual()

    def exportar_log_txt(self):
        try:
            conteudo = self.txt_log.get("1.0", "end").strip()
            if not conteudo:
                messagebox.showwarning("Exportar", "O log está vazio.")
                return
            caminho = filedialog.asksaveasfilename(
                title="Salvar log", defaultextension=".txt",
                filetypes=[("Texto", "*.txt")], initialfile="log_ssa_f8.txt")
            if not caminho:
                return
            with open(caminho, "w", encoding="utf-8") as f:
                f.write(conteudo)
            self.log(f"Log exportado: {caminho}")
            self.set_status("log exportado", self.COR_SUCESSO)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def salvar_layout_json(self):
        try:
            dados = {
                "usar_campo": self.cmb_usar.get(), "modo": self.cmb_modo.get(),
                "delay": self.entry_delay.get(), "delay_inicial": self.entry_delay_inicial.get(),
                "quantidade": self.entry_qtd.get(), "tabs_cdc": self.entry_tabs_cdc.get(),
                "tabs_final": self.entry_tabs_final.get(), "limpar_campo": self.cmb_limpar.get(),
                "acao_final": self.cmb_acao_final.get(), "linha_inicial": self.entry_linha_inicial.get(),
            }
            caminho = filedialog.asksaveasfilename(
                title="Salvar layout", defaultextension=".json",
                filetypes=[("JSON", "*.json")], initialfile="layout_ssa_f8.json")
            if not caminho:
                return
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=4)
            self.log(f"Layout salvo: {caminho}")
            self.set_status("layout salvo", self.COR_SUCESSO)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def carregar_layout_json(self):
        try:
            caminho = filedialog.askopenfilename(title="Carregar layout", filetypes=[("JSON", "*.json")])
            if not caminho:
                return
            with open(caminho, "r", encoding="utf-8") as f:
                d = json.load(f)
            self.cmb_usar.set(d.get("usar_campo", "VALOR"))
            self.cmb_modo.set(d.get("modo", "TAB"))
            for entry, key, default in [
                (self.entry_delay, "delay", "0,15"),
                (self.entry_delay_inicial, "delay_inicial", "3"),
                (self.entry_qtd, "quantidade", "TODOS"),
                (self.entry_tabs_cdc, "tabs_cdc", "3"),
                (self.entry_tabs_final, "tabs_final", "0"),
                (self.entry_linha_inicial, "linha_inicial", "1"),
            ]:
                entry.delete(0, "end")
                entry.insert(0, d.get(key, default))
            self.cmb_limpar.set(d.get("limpar_campo", "SIM"))
            self.cmb_acao_final.set(d.get("acao_final", "ENTER"))
            self.log(f"Layout carregado: {caminho}")
            self.set_status("layout carregado", self.COR_SUCESSO)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # =========================================================
    # FECHAMENTO
    # =========================================================
    def fechar_app(self):
        try:
            if self.hk_listener:
                self.hk_listener.stop()
        except Exception:
            pass
        self.app.destroy()

    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = PreenchedorSSAF8Simples()
    app.run()

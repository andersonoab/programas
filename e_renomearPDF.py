import customtkinter as ctk
from tkinter import filedialog, messagebox
import PyPDF2
import os
import re
import unicodedata
from datetime import datetime
import math

# ══════════════════════════════════════════════════════════════════
# SONOVA BRAND PALETTE — built from primary #0083CB
# ══════════════════════════════════════════════════════════════════

SONOVA = {
    "primary":         "#0083CB",
    "primary_dark":    "#006BA7",
    "primary_darker":  "#004F7C",
    "primary_light":   "#E6F3FB",
    "primary_soft":    "#B3D9F0",

    "bg":              "#F5F8FA",
    "bg_white":        "#FFFFFF",
    "bg_card":         "#FFFFFF",
    "bg_input":        "#F0F4F8",

    "text_primary":    "#0D1B2A",
    "text_secondary":  "#2C3E50",
    "text_muted":      "#546A7B",
    "text_on_primary": "#FFFFFF",

    "border":          "#D4DEE8",
    "border_light":    "#E8EFF5",

    "success":         "#0CAA6B",
    "success_hover":   "#0A8F5A",
    "warning":         "#E6A817",
    "danger":          "#D94040",

    "header_bg":       "#0083CB",
    "tab_active":      "#0083CB",
    "tab_inactive":    "#E0EBF5",
    "tab_text_active": "#FFFFFF",
    "tab_text_inact":  "#2C3E50",

    "footer_bg":       "#E6F3FB",
    "footer_text":     "#2C3E50",
}

ctk.set_appearance_mode("light")


class PDFRenamerPro:
    def __init__(self):
        self.app = ctk.CTk()
        self.app.title("Sonova | PDF Renamer & Splitter")
        self.app.geometry("1200x870")
        self.app.minsize(1050, 770)
        self.app.configure(fg_color=SONOVA["bg"])

        # State — Renamer
        self.files = []
        self.current_index = 0
        self.pdf_path = ""
        self.document_id = ""
        self.document_type = ""
        self.manual_name_var = ctk.StringVar()
        self.manual_name_var.trace_add("write", lambda *_: self.update_suggestion_label())

        # State — Splitter
        self.split_source = ""
        self.split_total_pages = 0

        self._build_ui()

    # ══════════════════════════════════════════════════════════
    # BUILD UI
    # ══════════════════════════════════════════════════════════
    def _build_ui(self):
        header = ctk.CTkFrame(self.app, height=72, fg_color=SONOVA["header_bg"], corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        header_inner = ctk.CTkFrame(header, fg_color="transparent")
        header_inner.pack(expand=True)

        title_block = ctk.CTkFrame(header_inner, fg_color="transparent")
        title_block.pack(side="left")

        ctk.CTkLabel(
            title_block,
            text="PDF Renamer & Splitter",
            font=("Segoe UI Semibold", 21),
            text_color="#FFFFFF"
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_block,
            text="Sonova — Ferramentas de RH",
            font=("Segoe UI", 11),
            text_color="#B3D9F0"
        ).pack(anchor="w")

        tab_bar = ctk.CTkFrame(self.app, height=50, fg_color=SONOVA["bg_white"], corner_radius=0)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)

        self.tab_rename_btn = ctk.CTkButton(
            tab_bar,
            text="Renomear PDFs",
            font=("Segoe UI Semibold", 13),
            height=34,
            corner_radius=6,
            fg_color=SONOVA["tab_active"],
            text_color=SONOVA["tab_text_active"],
            hover_color=SONOVA["primary_dark"],
            command=lambda: self._switch_tab("rename")
        )
        self.tab_rename_btn.pack(side="left", padx=(20, 4), pady=8)

        self.tab_split_btn = ctk.CTkButton(
            tab_bar,
            text="Separar PDF",
            font=("Segoe UI Semibold", 13),
            height=34,
            corner_radius=6,
            fg_color=SONOVA["tab_inactive"],
            text_color=SONOVA["tab_text_inact"],
            hover_color=SONOVA["primary_soft"],
            command=lambda: self._switch_tab("split")
        )
        self.tab_split_btn.pack(side="left", padx=4, pady=8)

        sep = ctk.CTkFrame(self.app, height=1, fg_color=SONOVA["border"], corner_radius=0)
        sep.pack(fill="x")

        self.content = ctk.CTkFrame(self.app, fg_color=SONOVA["bg"])
        self.content.pack(fill="both", expand=True, padx=0, pady=0)

        self.rename_page = ctk.CTkFrame(self.content, fg_color=SONOVA["bg"])
        self.split_page = ctk.CTkFrame(self.content, fg_color=SONOVA["bg"])

        self._build_rename_page()
        self._build_split_page()

        self.current_tab = "rename"
        self.rename_page.pack(fill="both", expand=True)

        footer = ctk.CTkFrame(self.app, height=32, fg_color=SONOVA["footer_bg"], corner_radius=0)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkLabel(
            footer,
            text="Anderson Marinho | Igarapé Digital",
            font=("Segoe UI", 11),
            text_color=SONOVA["footer_text"]
        ).pack(side="right", padx=18, pady=4)

    # ══════════════════════════════════════════════════════════
    # TAB SWITCHING
    # ══════════════════════════════════════════════════════════
    def _switch_tab(self, tab):
        if tab == self.current_tab:
            return

        self.current_tab = tab
        self.rename_page.pack_forget()
        self.split_page.pack_forget()

        if tab == "rename":
            self.tab_rename_btn.configure(
                fg_color=SONOVA["tab_active"],
                text_color=SONOVA["tab_text_active"]
            )
            self.tab_split_btn.configure(
                fg_color=SONOVA["tab_inactive"],
                text_color=SONOVA["tab_text_inact"]
            )
            self.rename_page.pack(fill="both", expand=True)
        else:
            self.tab_rename_btn.configure(
                fg_color=SONOVA["tab_inactive"],
                text_color=SONOVA["tab_text_inact"]
            )
            self.tab_split_btn.configure(
                fg_color=SONOVA["tab_active"],
                text_color=SONOVA["tab_text_active"]
            )
            self.split_page.pack(fill="both", expand=True)

    # ══════════════════════════════════════════════════════════
    # RENAME PAGE
    # ══════════════════════════════════════════════════════════
    def _build_rename_page(self):
        page = self.rename_page

        toolbar = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        toolbar.pack(fill="x", padx=18, pady=(18, 0))

        left = ctk.CTkFrame(toolbar, fg_color="transparent")
        left.pack(side="left", padx=12, pady=10)

        self.btn_folder = ctk.CTkButton(
            left,
            text="Selecionar Pasta",
            font=("Segoe UI Semibold", 13),
            width=190,
            height=38,
            fg_color=SONOVA["primary"],
            hover_color=SONOVA["primary_dark"],
            text_color="#FFFFFF",
            corner_radius=8,
            command=self.select_folder
        )
        self.btn_folder.pack(side="left", padx=(0, 8))

        self.btn_prev = ctk.CTkButton(
            left,
            text="Anterior",
            font=("Segoe UI", 13),
            width=120,
            height=38,
            fg_color=SONOVA["bg_input"],
            hover_color=SONOVA["primary_light"],
            text_color=SONOVA["text_secondary"],
            border_width=1,
            border_color=SONOVA["border"],
            corner_radius=8,
            command=self.previous_pdf
        )
        self.btn_prev.pack(side="left", padx=4)

        self.btn_skip = ctk.CTkButton(
            left,
            text="Pular",
            font=("Segoe UI", 13),
            width=100,
            height=38,
            fg_color=SONOVA["bg_input"],
            hover_color=SONOVA["primary_light"],
            text_color=SONOVA["text_secondary"],
            border_width=1,
            border_color=SONOVA["border"],
            corner_radius=8,
            command=self.skip_pdf
        )
        self.btn_skip.pack(side="left", padx=4)

        right = ctk.CTkFrame(toolbar, fg_color="transparent")
        right.pack(side="right", padx=12, pady=10)

        self.lbl_status = ctk.CTkLabel(
            right,
            text="0 / 0",
            font=("Segoe UI Semibold", 14),
            text_color=SONOVA["primary"],
            fg_color=SONOVA["primary_light"],
            corner_radius=6,
            width=80,
            height=34
        )
        self.lbl_status.pack(side="right")

        ctk.CTkLabel(
            right,
            text="Progresso:",
            font=("Segoe UI", 12),
            text_color=SONOVA["text_muted"]
        ).pack(side="right", padx=(0, 8))

        info_row = ctk.CTkFrame(page, fg_color="transparent")
        info_row.pack(fill="x", padx=18, pady=(12, 0))
        info_row.columnconfigure(0, weight=1)
        info_row.columnconfigure(1, weight=1)
        info_row.columnconfigure(2, weight=1)

        def _card(parent, col, label, padx_val):
            c = ctk.CTkFrame(
                parent,
                fg_color=SONOVA["bg_card"],
                corner_radius=8,
                height=62,
                border_width=1,
                border_color=SONOVA["border_light"]
            )
            c.grid(row=0, column=col, sticky="ew", padx=padx_val)
            c.pack_propagate(False)

            ctk.CTkLabel(
                c,
                text=label,
                font=("Segoe UI", 10),
                text_color=SONOVA["text_muted"]
            ).pack(anchor="w", padx=12, pady=(8, 0))

            lbl = ctk.CTkLabel(
                c,
                text="—",
                font=("Segoe UI Semibold", 12),
                text_color=SONOVA["text_primary"],
                anchor="w"
            )
            lbl.pack(anchor="w", padx=12, pady=(2, 8))
            return lbl

        self.lbl_file = _card(info_row, 0, "ARQUIVO", (0, 4))
        self.lbl_cpf = _card(info_row, 1, "CPF/PIS DETECTADO", (4, 4))
        self.lbl_cpf.configure(text_color=SONOVA["primary"])
        self.lbl_sugestao = _card(info_row, 2, "NOME SUGERIDO", (4, 0))
        self.lbl_sugestao.configure(text_color=SONOVA["success"])

        name_card = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=8,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        name_card.pack(fill="x", padx=18, pady=(12, 0))

        ctk.CTkLabel(
            name_card,
            text="Nome manual (opcional — se preenchido, ignora a seleção do texto):",
            font=("Segoe UI", 12),
            text_color=SONOVA["text_secondary"]
        ).pack(anchor="w", padx=14, pady=(10, 4))

        input_row = ctk.CTkFrame(name_card, fg_color="transparent")
        input_row.pack(fill="x", padx=14, pady=(0, 10))

        self.entry_name = ctk.CTkEntry(
            input_row,
            textvariable=self.manual_name_var,
            font=("Segoe UI", 13),
            height=38,
            fg_color=SONOVA["bg_input"],
            border_color=SONOVA["border"],
            text_color=SONOVA["text_primary"],
            placeholder_text="Ex: João da Silva Pereira",
            placeholder_text_color=SONOVA["text_muted"],
            corner_radius=8
        )
        self.entry_name.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_rename = ctk.CTkButton(
            input_row,
            text="Renomear e Próximo",
            font=("Segoe UI Semibold", 13),
            width=200,
            height=38,
            fg_color=SONOVA["success"],
            hover_color=SONOVA["success_hover"],
            text_color="#FFFFFF",
            corner_radius=8,
            command=self.rename_selected
        )
        self.btn_rename.pack(side="right")

        self.entry_name.bind("<Return>", lambda e: self.rename_selected())

        text_frame = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        text_frame.pack(fill="both", expand=True, padx=18, pady=(12, 0))

        ctk.CTkLabel(
            text_frame,
            text="CONTEÚDO DO PDF — selecione um trecho para usar como nome",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w", padx=14, pady=(10, 0))

        self.textbox = ctk.CTkTextbox(
            text_frame,
            font=("Consolas", 13),
            fg_color=SONOVA["bg_input"],
            text_color=SONOVA["text_primary"],
            border_width=1,
            border_color=SONOVA["border"],
            corner_radius=6
        )
        self.textbox.pack(fill="both", expand=True, padx=12, pady=(6, 12))

        log_frame = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        log_frame.pack(fill="x", padx=18, pady=(12, 18))

        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.pack(fill="x", padx=14, pady=(10, 0))

        ctk.CTkLabel(
            log_header,
            text="LOG DE ATIVIDADE",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(side="left")

        ctk.CTkLabel(
            log_header,
            text="Enter = renomear | Shift+Enter = pular | Selecione texto ou digite o nome",
            font=("Segoe UI", 11),
            text_color=SONOVA["text_muted"]
        ).pack(side="right")

        self.log = ctk.CTkTextbox(
            log_frame,
            height=110,
            font=("Consolas", 12),
            fg_color=SONOVA["bg_input"],
            text_color=SONOVA["text_secondary"],
            border_width=0,
            corner_radius=6
        )
        self.log.pack(fill="x", padx=12, pady=(6, 12))

        self.textbox.bind("<Return>", self.enter_renomear)
        self.textbox.bind("<Shift-Return>", self.shift_enter_skip)
        self.textbox.bind("<ButtonRelease-1>", self.on_text_selection)
        self.textbox.bind("<KeyRelease>", self.on_text_selection)

        self.update_navigation_buttons()

    # ══════════════════════════════════════════════════════════
    # SPLIT PAGE
    # ══════════════════════════════════════════════════════════
    def _build_split_page(self):
        page = self.split_page

        top_card = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        top_card.pack(fill="x", padx=18, pady=(18, 0))

        ctk.CTkLabel(
            top_card,
            text="Separar PDF em partes por quantidade de páginas",
            font=("Segoe UI Semibold", 16),
            text_color=SONOVA["text_primary"]
        ).pack(anchor="w", padx=16, pady=(16, 4))

        ctk.CTkLabel(
            top_card,
            text="Selecione um PDF, defina quantas páginas cada parte terá e clique em Separar.",
            font=("Segoe UI", 13),
            text_color=SONOVA["text_secondary"]
        ).pack(anchor="w", padx=16, pady=(0, 16))

        src_card = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        src_card.pack(fill="x", padx=18, pady=(12, 0))

        ctk.CTkLabel(
            src_card,
            text="ARQUIVO DE ORIGEM",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w", padx=16, pady=(12, 4))

        file_row = ctk.CTkFrame(src_card, fg_color="transparent")
        file_row.pack(fill="x", padx=16, pady=(0, 12))

        self.lbl_split_file = ctk.CTkLabel(
            file_row,
            text="Nenhum arquivo selecionado",
            font=("Segoe UI", 13),
            text_color=SONOVA["text_secondary"],
            anchor="w"
        )
        self.lbl_split_file.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            file_row,
            text="Selecionar PDF",
            font=("Segoe UI Semibold", 13),
            width=180,
            height=38,
            fg_color=SONOVA["primary"],
            hover_color=SONOVA["primary_dark"],
            text_color="#FFFFFF",
            corner_radius=8,
            command=self.select_split_file
        ).pack(side="right")

        cfg_card = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        cfg_card.pack(fill="x", padx=18, pady=(12, 0))

        info_inner = ctk.CTkFrame(cfg_card, fg_color="transparent")
        info_inner.pack(fill="x", padx=16, pady=14)

        tp = ctk.CTkFrame(info_inner, fg_color="transparent")
        tp.pack(side="left", padx=(0, 30))

        ctk.CTkLabel(
            tp,
            text="TOTAL DE PÁGINAS",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w")

        self.lbl_total_pages = ctk.CTkLabel(
            tp,
            text="—",
            font=("Segoe UI Semibold", 20),
            text_color=SONOVA["text_primary"]
        )
        self.lbl_total_pages.pack(anchor="w")

        pps = ctk.CTkFrame(info_inner, fg_color="transparent")
        pps.pack(side="left", padx=(0, 30))

        ctk.CTkLabel(
            pps,
            text="PÁGINAS POR ARQUIVO",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w")

        self.split_pages_var = ctk.StringVar(value="1")
        self.entry_pages = ctk.CTkEntry(
            pps,
            textvariable=self.split_pages_var,
            font=("Segoe UI Semibold", 18),
            width=80,
            height=40,
            fg_color=SONOVA["bg_input"],
            border_color=SONOVA["border"],
            text_color=SONOVA["text_primary"],
            corner_radius=8,
            justify="center"
        )
        self.entry_pages.pack(anchor="w", pady=(2, 0))

        rc = ctk.CTkFrame(info_inner, fg_color="transparent")
        rc.pack(side="left", padx=(0, 30))

        ctk.CTkLabel(
            rc,
            text="ARQUIVOS GERADOS",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w")

        self.lbl_result_count = ctk.CTkLabel(
            rc,
            text="—",
            font=("Segoe UI Semibold", 20),
            text_color=SONOVA["primary"]
        )
        self.lbl_result_count.pack(anchor="w")

        of = ctk.CTkFrame(info_inner, fg_color="transparent")
        of.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            of,
            text="PASTA DE SAÍDA",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w")

        self.lbl_output_folder = ctk.CTkLabel(
            of,
            text="Mesma pasta do original",
            font=("Segoe UI", 12),
            text_color=SONOVA["text_secondary"]
        )
        self.lbl_output_folder.pack(anchor="w")

        self.split_pages_var.trace_add("write", lambda *_: self._update_split_preview())

        naming_card = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        naming_card.pack(fill="x", padx=18, pady=(12, 0))

        ctk.CTkLabel(
            naming_card,
            text="PREFIXO DO NOME DOS ARQUIVOS (opcional)",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w", padx=16, pady=(12, 4))

        self.split_prefix_var = ctk.StringVar()
        ctk.CTkEntry(
            naming_card,
            textvariable=self.split_prefix_var,
            font=("Segoe UI", 13),
            height=38,
            fg_color=SONOVA["bg_input"],
            border_color=SONOVA["border"],
            text_color=SONOVA["text_primary"],
            placeholder_text="Ex: Contracheque_Marco",
            placeholder_text_color=SONOVA["text_muted"],
            corner_radius=8
        ).pack(fill="x", padx=16, pady=(0, 12))

        action_row = ctk.CTkFrame(page, fg_color="transparent")
        action_row.pack(fill="x", padx=18, pady=(16, 0))

        self.btn_split = ctk.CTkButton(
            action_row,
            text="Separar PDF",
            font=("Segoe UI Semibold", 15),
            width=220,
            height=46,
            fg_color=SONOVA["primary"],
            hover_color=SONOVA["primary_dark"],
            text_color="#FFFFFF",
            corner_radius=10,
            command=self.execute_split
        )
        self.btn_split.pack(anchor="center")

        slog_frame = ctk.CTkFrame(
            page,
            fg_color=SONOVA["bg_card"],
            corner_radius=10,
            border_width=1,
            border_color=SONOVA["border_light"]
        )
        slog_frame.pack(fill="both", expand=True, padx=18, pady=(12, 18))

        ctk.CTkLabel(
            slog_frame,
            text="LOG",
            font=("Segoe UI", 10),
            text_color=SONOVA["text_muted"]
        ).pack(anchor="w", padx=14, pady=(10, 0))

        self.split_log = ctk.CTkTextbox(
            slog_frame,
            font=("Consolas", 12),
            fg_color=SONOVA["bg_input"],
            text_color=SONOVA["text_secondary"],
            border_width=0,
            corner_radius=6
        )
        self.split_log.pack(fill="both", expand=True, padx=12, pady=(6, 12))

    # ══════════════════════════════════════════════════════════
    # UTILS
    # ══════════════════════════════════════════════════════════
    def write_log(self, text):
        now = datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{now}] {text}\n")
        self.log.see("end")

    def write_split_log(self, text):
        now = datetime.now().strftime("%H:%M:%S")
        self.split_log.insert("end", f"[{now}] {text}\n")
        self.split_log.see("end")

    def normalize_name(self, name):
        name = unicodedata.normalize("NFKD", name)
        name = "".join(c for c in name if not unicodedata.combining(c))
        name = name.upper()
        name = re.sub(r"[^\w\s]", "", name)
        name = re.sub(r"\s+", "_", name)
        name = re.sub(r"_+", "_", name)
        return name.strip("_")

    def clean_selected_name(self, text):
        text = text.strip()

        substituicoes = [
            "Nome:",
            "NOME:",
            "Nome",
            "NOME",
            "CPF:",
            "CPF",
            "PIS:",
            "PIS",
            "PASEP:",
            "PASEP",
            "NIT:",
            "NIT",
            "CNPJ:",
            "CNPJ",
            "Beneficiário:",
            "BENEFICIÁRIO:",
            "BENEFICIARIO:",
            "BENEFICIÁRIO",
            "BENEFICIARIO"
        ]

        for item in substituicoes:
            text = text.replace(item, "")

        text = text.strip(" :-\n\t")
        text = re.sub(r"\s+", " ", text)

        return text.strip()

    def get_selected_text(self):
        try:
            ranges = self.textbox.tag_ranges("sel")
            if not ranges:
                return ""
            return self.textbox.get(ranges[0], ranges[1]).strip()
        except Exception:
            return ""

    def update_status_label(self):
        total = len(self.files)
        atual = self.current_index + 1 if total > 0 and self.current_index < total else total
        self.lbl_status.configure(text=f"{atual} / {total}")

    def update_file_info(self):
        if self.pdf_path:
            self.lbl_file.configure(text=os.path.basename(self.pdf_path))
        else:
            self.lbl_file.configure(text="—")

        if self.document_id and self.document_id != "SEMID":
            if self.document_type == "CPF":
                doc_fmt = f"{self.document_id[:3]}.{self.document_id[3:6]}.{self.document_id[6:9]}-{self.document_id[9:]}"
                self.lbl_cpf.configure(text=f"CPF: {doc_fmt}")
            elif self.document_type == "PIS":
                doc_fmt = f"{self.document_id[:3]}.{self.document_id[3:8]}.{self.document_id[8:10]}-{self.document_id[10:]}"
                self.lbl_cpf.configure(text=f"PIS: {doc_fmt}")
            else:
                self.lbl_cpf.configure(text=self.document_id)
        else:
            self.lbl_cpf.configure(text="Não encontrado")

    def update_navigation_buttons(self):
        tem = len(self.files) > 0
        ant = tem and self.current_index > 0
        atu = tem and self.current_index < len(self.files)

        self.btn_prev.configure(state="normal" if ant else "disabled")
        self.btn_skip.configure(state="normal" if atu else "disabled")
        self.btn_rename.configure(state="normal" if atu else "disabled")

    def update_suggestion_label(self):
        manual = self.manual_name_var.get().strip()

        if manual:
            nome_final = self.normalize_name(manual)
            if nome_final:
                doc_part = self.build_document_prefix()
                self.lbl_sugestao.configure(text=f"{doc_part}_{nome_final}.pdf")
                return

        selecionado = self.get_selected_text()

        if not selecionado:
            self.lbl_sugestao.configure(text="—")
            return

        nome_limpo = self.clean_selected_name(selecionado)
        nome_final = self.normalize_name(nome_limpo)

        if not nome_final:
            self.lbl_sugestao.configure(text="—")
            return

        doc_part = self.build_document_prefix()
        self.lbl_sugestao.configure(text=f"{doc_part}_{nome_final}.pdf")

    def build_document_prefix(self):
        if self.document_id and self.document_id != "SEMID":
            return f"{self.document_type}_{self.document_id}"

        return "SEMID"

    # ══════════════════════════════════════════════════════════
    # EVENTS
    # ══════════════════════════════════════════════════════════
    def on_text_selection(self, event=None):
        self.update_suggestion_label()

    def enter_renomear(self, event):
        self.rename_selected()
        return "break"

    def shift_enter_skip(self, event):
        self.skip_pdf()
        return "break"

    # ══════════════════════════════════════════════════════════
    # FOLDER / PDF
    # ══════════════════════════════════════════════════════════
    def select_folder(self):
        folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")

        if not folder:
            return

        self.files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ]

        self.files.sort(key=lambda x: os.path.basename(x).lower())
        self.current_index = 0
        self.pdf_path = ""
        self.document_id = ""
        self.document_type = ""
        self.manual_name_var.set("")

        self.write_log(f"{len(self.files)} PDFs encontrados em: {folder}")

        if not self.files:
            self.textbox.delete("1.0", "end")
            self.lbl_file.configure(text="—")
            self.lbl_cpf.configure(text="—")
            self.lbl_sugestao.configure(text="—")
            self.update_status_label()
            self.update_navigation_buttons()
            messagebox.showwarning("Aviso", "Nenhum PDF encontrado na pasta selecionada.")
            return

        self.load_pdf()

    def extract_text_from_pdf(self, pdf_path):
        text = ""

        with open(pdf_path, "rb") as f:
            pdf = PyPDF2.PdfReader(f)

            for page in pdf.pages:
                try:
                    text += page.extract_text() or ""
                    text += "\n"
                except Exception:
                    pass

        return text

    # ══════════════════════════════════════════════════════════
    # VALIDADORES CPF / PIS / CNPJ
    # ══════════════════════════════════════════════════════════
    def validate_cpf(self, cpf):
        cpf = re.sub(r"\D", "", cpf)

        if len(cpf) != 11:
            return False

        if cpf == cpf[0] * 11:
            return False

        soma = 0
        for i in range(9):
            soma += int(cpf[i]) * (10 - i)

        digito_1 = 11 - (soma % 11)

        if digito_1 >= 10:
            digito_1 = 0

        if digito_1 != int(cpf[9]):
            return False

        soma = 0
        for i in range(10):
            soma += int(cpf[i]) * (11 - i)

        digito_2 = 11 - (soma % 11)

        if digito_2 >= 10:
            digito_2 = 0

        if digito_2 != int(cpf[10]):
            return False

        return True

    def validate_pis(self, pis):
        pis = re.sub(r"\D", "", pis)

        if len(pis) != 11:
            return False

        if pis == pis[0] * 11:
            return False

        pesos = [3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

        soma = 0
        for i in range(10):
            soma += int(pis[i]) * pesos[i]

        resto = soma % 11
        digito = 11 - resto

        if digito == 10 or digito == 11:
            digito = 0

        if digito != int(pis[10]):
            return False

        return True

    def validate_cnpj(self, cnpj):
        cnpj = re.sub(r"\D", "", cnpj)

        if len(cnpj) != 14:
            return False

        if cnpj == cnpj[0] * 14:
            return False

        pesos_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        pesos_2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]

        soma = 0
        for i in range(12):
            soma += int(cnpj[i]) * pesos_1[i]

        resto = soma % 11
        digito_1 = 0 if resto < 2 else 11 - resto

        if digito_1 != int(cnpj[12]):
            return False

        soma = 0
        for i in range(13):
            soma += int(cnpj[i]) * pesos_2[i]

        resto = soma % 11
        digito_2 = 0 if resto < 2 else 11 - resto

        if digito_2 != int(cnpj[13]):
            return False

        return True

    # ══════════════════════════════════════════════════════════
    # EXTRAÇÃO CPF/PIS SEM PEGAR CNPJ
    # ══════════════════════════════════════════════════════════
    def find_document_in_text(self, text):
        texto = text or ""

        texto_sem_cnpj = self.remove_cnpj_from_text(texto)

        candidatos_contexto = []
        candidatos_gerais = []

        padroes_contexto = [
            r"\bCPF\b[:\s\-]*([0-9][0-9\.\-\s/]{8,25}[0-9])",
            r"\bPIS\b[:\s\-]*([0-9][0-9\.\-\s/]{8,25}[0-9])",
            r"\bPASEP\b[:\s\-]*([0-9][0-9\.\-\s/]{8,25}[0-9])",
            r"\bNIT\b[:\s\-]*([0-9][0-9\.\-\s/]{8,25}[0-9])"
        ]

        for padrao in padroes_contexto:
            encontrados = re.findall(padrao, texto_sem_cnpj, flags=re.IGNORECASE)

            for item in encontrados:
                numero = re.sub(r"\D", "", item)

                if len(numero) == 11:
                    candidatos_contexto.append(numero)

        padroes_gerais = [
            r"\b\d{3}\.\d{3}\.\d{3}[\-/]\d{2}\b",
            r"\b\d{3}\.\d{3}\.\d{5}\b",
            r"\b\d{3}\s+\d{3}\s+\d{3}\s+\d{2}\b",
            r"\b\d{3}\-\d{3}\-\d{3}\-\d{2}\b",
            r"\b\d{3}\.\d{5}\.\d{2}-\d{1}\b",
            r"(?<!\d)\d{11}(?!\d)",
            r"(?<!\d)\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d[\.\-\s/]*\d(?!\d)"
        ]

        for padrao in padroes_gerais:
            encontrados = re.findall(padrao, texto_sem_cnpj, flags=re.IGNORECASE)

            for item in encontrados:
                numero = re.sub(r"\D", "", item)

                if len(numero) == 11:
                    candidatos_gerais.append(numero)

        candidatos_contexto = self.unique_list(candidatos_contexto)
        candidatos_gerais = self.unique_list(candidatos_gerais)

        for numero in candidatos_contexto:
            if self.validate_cpf(numero):
                return "CPF", numero

        for numero in candidatos_contexto:
            if self.validate_pis(numero):
                return "PIS", numero

        for numero in candidatos_gerais:
            if self.validate_cpf(numero):
                return "CPF", numero

        for numero in candidatos_gerais:
            if self.validate_pis(numero):
                return "PIS", numero

        return "", "SEMID"

    def remove_cnpj_from_text(self, text):
        texto = text or ""

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

    def unique_list(self, values):
        result = []

        for value in values:
            if value not in result:
                result.append(value)

        return result

    def load_pdf(self):
        self.manual_name_var.set("")

        if self.current_index >= len(self.files):
            self.textbox.delete("1.0", "end")
            self.pdf_path = ""
            self.document_id = ""
            self.document_type = ""
            self.update_status_label()
            self.update_file_info()
            self.update_suggestion_label()
            self.update_navigation_buttons()
            self.write_log("Todos os PDFs foram processados")
            messagebox.showinfo("Concluído", "Todos os PDFs foram processados.")
            return

        self.pdf_path = self.files[self.current_index]

        self.write_log(
            f"Abrindo PDF {self.current_index + 1}/{len(self.files)} — {os.path.basename(self.pdf_path)}"
        )

        try:
            text = self.extract_text_from_pdf(self.pdf_path)
        except Exception as e:
            text = ""
            self.write_log(f"Erro leitura PDF: {e}")

        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", text)

        self.document_type, self.document_id = self.find_document_in_text(text)

        if self.document_id != "SEMID":
            self.write_log(f"{self.document_type} encontrado e validado: {self.document_id}")
        else:
            self.write_log("CPF/PIS não encontrado. CNPJ foi ignorado.")

        self.update_status_label()
        self.update_file_info()
        self.update_suggestion_label()
        self.update_navigation_buttons()
        self.textbox.focus_set()

    # ══════════════════════════════════════════════════════════
    # NAV
    # ══════════════════════════════════════════════════════════
    def previous_pdf(self):
        if not self.files:
            self.write_log("Nenhum PDF carregado")
            return

        if self.current_index <= 0:
            self.write_log("Já está no primeiro PDF")
            return

        self.current_index -= 1
        self.load_pdf()

    def skip_pdf(self):
        if not self.pdf_path:
            self.write_log("Nenhum PDF carregado")
            return

        self.write_log(f"Pulado: {os.path.basename(self.pdf_path)}")
        self.current_index += 1
        self.load_pdf()

    # ══════════════════════════════════════════════════════════
    # RENAME
    # ══════════════════════════════════════════════════════════
    def build_new_name(self, nome_base):
        folder = os.path.dirname(self.pdf_path)
        doc_part = self.build_document_prefix()

        nome_arquivo = f"{doc_part}_{nome_base}.pdf"
        novo_caminho = os.path.join(folder, nome_arquivo)

        contador = 1

        while os.path.exists(novo_caminho) and os.path.abspath(novo_caminho) != os.path.abspath(self.pdf_path):
            nome_arquivo = f"{doc_part}_{nome_base}_{contador}.pdf"
            novo_caminho = os.path.join(folder, nome_arquivo)
            contador += 1

        return nome_arquivo, novo_caminho

    def rename_selected(self):
        if not self.pdf_path:
            self.write_log("Nenhum PDF carregado")
            return

        manual = self.manual_name_var.get().strip()

        if manual:
            nome_limpo = manual
        else:
            selected_text = self.get_selected_text()

            if not selected_text:
                self.write_log("Digite um nome ou selecione texto no PDF")
                return

            nome_limpo = self.clean_selected_name(selected_text)

        if not nome_limpo:
            self.write_log("Nome vazio — selecione ou digite um nome válido")
            return

        nome_normalizado = self.normalize_name(nome_limpo)

        if not nome_normalizado:
            self.write_log("Nome inválido após normalização")
            return

        novo_nome, novo_path = self.build_new_name(nome_normalizado)

        try:
            antigo_nome = os.path.basename(self.pdf_path)
            os.rename(self.pdf_path, novo_path)

            self.write_log(f"Renomeado: {antigo_nome} -> {novo_nome}")

            self.files[self.current_index] = novo_path
            self.current_index += 1
            self.load_pdf()

        except Exception as e:
            self.write_log(f"Erro renomear: {e}")
            messagebox.showerror("Erro", f"Erro ao renomear o arquivo:\n{e}")

    # ══════════════════════════════════════════════════════════
    # SPLIT
    # ══════════════════════════════════════════════════════════
    def select_split_file(self):
        path = filedialog.askopenfilename(
            title="Selecione o PDF para separar",
            filetypes=[("PDF", "*.pdf")]
        )

        if not path:
            return

        self.split_source = path
        self.lbl_split_file.configure(text=os.path.basename(path))

        try:
            with open(path, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                self.split_total_pages = len(reader.pages)
        except Exception as e:
            self.split_total_pages = 0
            self.write_split_log(f"Erro ao ler PDF: {e}")

        self.lbl_total_pages.configure(text=str(self.split_total_pages))
        self._update_split_preview()
        self.write_split_log(
            f"PDF selecionado: {os.path.basename(path)} — {self.split_total_pages} páginas"
        )

    def _update_split_preview(self):
        if self.split_total_pages == 0:
            self.lbl_result_count.configure(text="—")
            return

        try:
            pp = int(self.split_pages_var.get())

            if pp < 1:
                pp = 1

        except (ValueError, TypeError):
            self.lbl_result_count.configure(text="—")
            return

        self.lbl_result_count.configure(text=str(math.ceil(self.split_total_pages / pp)))

    def execute_split(self):
        if not self.split_source or not os.path.isfile(self.split_source):
            messagebox.showwarning("Aviso", "Selecione um PDF primeiro.")
            return

        try:
            pages_per = int(self.split_pages_var.get())

            if pages_per < 1:
                raise ValueError

        except (ValueError, TypeError):
            messagebox.showwarning("Aviso", "Informe um número válido de páginas por arquivo.")
            return

        prefix = self.split_prefix_var.get().strip()

        if not prefix:
            base_name = os.path.splitext(os.path.basename(self.split_source))[0]
            prefix = base_name

        prefix_safe = self.normalize_name(prefix) if prefix else "PDF"

        folder = os.path.dirname(self.split_source)
        output_dir = os.path.join(folder, f"{prefix_safe}_separado")
        os.makedirs(output_dir, exist_ok=True)

        try:
            with open(self.split_source, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                total = len(reader.pages)
                part = 1

                for start in range(0, total, pages_per):
                    end = min(start + pages_per, total)
                    writer = PyPDF2.PdfWriter()

                    for i in range(start, end):
                        writer.add_page(reader.pages[i])

                    out_name = f"{prefix_safe}_parte{part}.pdf"
                    out_path = os.path.join(output_dir, out_name)

                    with open(out_path, "wb") as out_f:
                        writer.write(out_f)

                    self.write_split_log(f"Criado: {out_name} páginas {start + 1} até {end}")
                    part += 1

            self.write_split_log(f"Concluído — {part - 1} arquivos em: {output_dir}")
            messagebox.showinfo("Concluído", f"{part - 1} arquivos criados em:\n{output_dir}")

        except Exception as e:
            self.write_split_log(f"Erro na separação: {e}")
            messagebox.showerror("Erro", f"Erro ao separar o PDF:\n{e}")

    # ══════════════════════════════════════════════════════════
    # RUN
    # ══════════════════════════════════════════════════════════
    def run(self):
        self.app.mainloop()


if __name__ == "__main__":
    app = PDFRenamerPro()
    app.run()

import os
import shutil
import threading
import traceback
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import customtkinter as ctk
import pandas as pd

try:
    import PyPDF2
    PYPDF2_DISPONIVEL = True
except Exception:
    PYPDF2_DISPONIVEL = False


APP_NAME = "Buscador OneDrive PDF - Igarapé Digital"
APP_VERSION = "1.1"

COR_AZUL = "#121C4E"
COR_AZUL_2 = "#003C64"
COR_AZUL_CLARO = "#0083CA"
COR_FUNDO = "#F5F7FA"
COR_CARD = "#FFFFFF"
COR_TEXTO = "#1F2937"
COR_TEXTO_SEC = "#6B7280"
COR_BORDA = "#D1D5DB"
COR_SUCESSO = "#166534"
COR_ALERTA = "#92400E"
COR_ERRO = "#991B1B"


class BuscadorOneDrivePDF(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(APP_NAME)
        self.geometry("1250x760")
        self.minsize(1100, 680)

        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.resultados = []
        self.busca_em_execucao = False
        self.copia_em_execucao = False
        self.thread_busca = None
        self.thread_copia = None

        self.var_pasta = tk.StringVar()
        self.var_pasta_destino = tk.StringVar()
        self.var_termo = tk.StringVar()
        self.var_extensoes = tk.StringVar(value=".pdf")
        self.var_buscar_nome = tk.BooleanVar(value=True)
        self.var_buscar_conteudo_pdf = tk.BooleanVar(value=False)
        self.var_incluir_subpastas = tk.BooleanVar(value=True)
        self.var_ignorar_temporarios = tk.BooleanVar(value=True)
        self.var_limite_resultados = tk.StringVar(value="1000")
        self.var_status = tk.StringVar(value="Pronto para buscar.")
        self.var_progresso = tk.DoubleVar(value=0)

        self._montar_layout()

    def _montar_layout(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.header = ctk.CTkFrame(self, fg_color=COR_AZUL, corner_radius=0, height=72)
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_columnconfigure(0, weight=1)

        self.lbl_titulo = ctk.CTkLabel(
            self.header,
            text="Buscador OneDrive PDF",
            font=("Segoe UI", 24, "bold"),
            text_color="white"
        )
        self.lbl_titulo.grid(row=0, column=0, sticky="w", padx=24, pady=(12, 0))

        self.lbl_subtitulo = ctk.CTkLabel(
            self.header,
            text="Busca segura por arquivos, conteúdo de PDFs e cópia dos resultados sem usar a pesquisa do Windows",
            font=("Segoe UI", 13),
            text_color="#E5E7EB"
        )
        self.lbl_subtitulo.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 12))

        self.frame_config = ctk.CTkFrame(self, fg_color=COR_CARD, corner_radius=14)
        self.frame_config.grid(row=1, column=0, sticky="ew", padx=18, pady=14)
        self.frame_config.grid_columnconfigure(1, weight=1)

        self.lbl_pasta = ctk.CTkLabel(
            self.frame_config,
            text="Pasta base:",
            font=("Segoe UI", 13, "bold"),
            text_color=COR_TEXTO
        )
        self.lbl_pasta.grid(row=0, column=0, sticky="w", padx=16, pady=(16, 8))

        self.entry_pasta = ctk.CTkEntry(
            self.frame_config,
            textvariable=self.var_pasta,
            height=36,
            placeholder_text="Selecione a pasta do OneDrive ou cole o caminho aqui"
        )
        self.entry_pasta.grid(row=0, column=1, sticky="ew", padx=8, pady=(16, 8))

        self.btn_pasta = ctk.CTkButton(
            self.frame_config,
            text="Selecionar pasta",
            fg_color=COR_AZUL_CLARO,
            hover_color=COR_AZUL_2,
            command=self.selecionar_pasta,
            height=36
        )
        self.btn_pasta.grid(row=0, column=2, sticky="e", padx=16, pady=(16, 8))

        self.lbl_destino = ctk.CTkLabel(
            self.frame_config,
            text="Pasta destino:",
            font=("Segoe UI", 13, "bold"),
            text_color=COR_TEXTO
        )
        self.lbl_destino.grid(row=1, column=0, sticky="w", padx=16, pady=8)

        self.entry_destino = ctk.CTkEntry(
            self.frame_config,
            textvariable=self.var_pasta_destino,
            height=36,
            placeholder_text="Selecione a pasta para copiar os arquivos encontrados"
        )
        self.entry_destino.grid(row=1, column=1, sticky="ew", padx=8, pady=8)

        self.btn_destino = ctk.CTkButton(
            self.frame_config,
            text="Selecionar destino",
            fg_color=COR_AZUL_CLARO,
            hover_color=COR_AZUL_2,
            command=self.selecionar_pasta_destino,
            height=36
        )
        self.btn_destino.grid(row=1, column=2, sticky="e", padx=16, pady=8)

        self.lbl_termo = ctk.CTkLabel(
            self.frame_config,
            text="Termo:",
            font=("Segoe UI", 13, "bold"),
            text_color=COR_TEXTO
        )
        self.lbl_termo.grid(row=2, column=0, sticky="w", padx=16, pady=8)

        self.entry_termo = ctk.CTkEntry(
            self.frame_config,
            textvariable=self.var_termo,
            height=36,
            placeholder_text="Exemplo: CPF, PIS, matrícula, holerite, admissão, nome do colaborador"
        )
        self.entry_termo.grid(row=2, column=1, sticky="ew", padx=8, pady=8)

        self.btn_buscar = ctk.CTkButton(
            self.frame_config,
            text="Buscar",
            fg_color=COR_AZUL,
            hover_color=COR_AZUL_2,
            command=self.iniciar_busca,
            height=36
        )
        self.btn_buscar.grid(row=2, column=2, sticky="e", padx=16, pady=8)

        self.lbl_extensoes = ctk.CTkLabel(
            self.frame_config,
            text="Extensões:",
            font=("Segoe UI", 13, "bold"),
            text_color=COR_TEXTO
        )
        self.lbl_extensoes.grid(row=3, column=0, sticky="w", padx=16, pady=8)

        self.entry_extensoes = ctk.CTkEntry(
            self.frame_config,
            textvariable=self.var_extensoes,
            height=36,
            placeholder_text="Exemplo: .pdf ou .pdf;.xlsx;.docx"
        )
        self.entry_extensoes.grid(row=3, column=1, sticky="ew", padx=8, pady=8)

        self.btn_limpar = ctk.CTkButton(
            self.frame_config,
            text="Limpar",
            fg_color="#6B7280",
            hover_color="#4B5563",
            command=self.limpar_resultados,
            height=36
        )
        self.btn_limpar.grid(row=3, column=2, sticky="e", padx=16, pady=8)

        self.frame_opcoes = ctk.CTkFrame(self.frame_config, fg_color="#F9FAFB", corner_radius=10)
        self.frame_opcoes.grid(row=4, column=0, columnspan=3, sticky="ew", padx=16, pady=(8, 16))
        self.frame_opcoes.grid_columnconfigure(6, weight=1)

        self.chk_nome = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Buscar no nome do arquivo",
            variable=self.var_buscar_nome,
            text_color=COR_TEXTO,
            fg_color=COR_AZUL_CLARO
        )
        self.chk_nome.grid(row=0, column=0, sticky="w", padx=12, pady=12)

        self.chk_pdf = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Buscar dentro do PDF",
            variable=self.var_buscar_conteudo_pdf,
            text_color=COR_TEXTO,
            fg_color=COR_AZUL_CLARO
        )
        self.chk_pdf.grid(row=0, column=1, sticky="w", padx=12, pady=12)

        self.chk_subpastas = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Incluir subpastas",
            variable=self.var_incluir_subpastas,
            text_color=COR_TEXTO,
            fg_color=COR_AZUL_CLARO
        )
        self.chk_subpastas.grid(row=0, column=2, sticky="w", padx=12, pady=12)

        self.chk_temp = ctk.CTkCheckBox(
            self.frame_opcoes,
            text="Ignorar temporários",
            variable=self.var_ignorar_temporarios,
            text_color=COR_TEXTO,
            fg_color=COR_AZUL_CLARO
        )
        self.chk_temp.grid(row=0, column=3, sticky="w", padx=12, pady=12)

        self.lbl_limite = ctk.CTkLabel(
            self.frame_opcoes,
            text="Limite:",
            text_color=COR_TEXTO
        )
        self.lbl_limite.grid(row=0, column=4, sticky="e", padx=(20, 6), pady=12)

        self.entry_limite = ctk.CTkEntry(
            self.frame_opcoes,
            textvariable=self.var_limite_resultados,
            width=80,
            height=30
        )
        self.entry_limite.grid(row=0, column=5, sticky="w", padx=6, pady=12)

        self.frame_resultados = ctk.CTkFrame(self, fg_color=COR_CARD, corner_radius=14)
        self.frame_resultados.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 14))
        self.frame_resultados.grid_columnconfigure(0, weight=1)
        self.frame_resultados.grid_rowconfigure(1, weight=1)

        self.frame_barra = ctk.CTkFrame(self.frame_resultados, fg_color=COR_CARD)
        self.frame_barra.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 4))
        self.frame_barra.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            self.frame_barra,
            textvariable=self.var_status,
            text_color=COR_TEXTO_SEC,
            font=("Segoe UI", 12)
        )
        self.lbl_status.grid(row=0, column=0, sticky="w", padx=4)

        self.progress = ctk.CTkProgressBar(
            self.frame_barra,
            variable=self.var_progresso,
            height=10,
            progress_color=COR_AZUL_CLARO
        )
        self.progress.grid(row=1, column=0, sticky="ew", padx=4, pady=(6, 4))
        self.progress.set(0)

        self.frame_tree = tk.Frame(self.frame_resultados, bg=COR_CARD)
        self.frame_tree.grid(row=1, column=0, sticky="nsew", padx=12, pady=8)
        self.frame_tree.grid_columnconfigure(0, weight=1)
        self.frame_tree.grid_rowconfigure(0, weight=1)

        self.colunas = (
            "tipo",
            "arquivo",
            "extensao",
            "pasta",
            "pagina",
            "tamanho",
            "modificado",
            "caminho"
        )

        self.tree = ttk.Treeview(
            self.frame_tree,
            columns=self.colunas,
            show="headings",
            height=16,
            selectmode="extended"
        )

        self.tree.heading("tipo", text="Tipo")
        self.tree.heading("arquivo", text="Arquivo")
        self.tree.heading("extensao", text="Ext.")
        self.tree.heading("pasta", text="Pasta")
        self.tree.heading("pagina", text="Página")
        self.tree.heading("tamanho", text="Tamanho KB")
        self.tree.heading("modificado", text="Modificado")
        self.tree.heading("caminho", text="Caminho completo")

        self.tree.column("tipo", width=130, anchor="w")
        self.tree.column("arquivo", width=250, anchor="w")
        self.tree.column("extensao", width=70, anchor="center")
        self.tree.column("pasta", width=260, anchor="w")
        self.tree.column("pagina", width=70, anchor="center")
        self.tree.column("tamanho", width=90, anchor="e")
        self.tree.column("modificado", width=140, anchor="center")
        self.tree.column("caminho", width=460, anchor="w")

        self.scroll_y = ttk.Scrollbar(self.frame_tree, orient="vertical", command=self.tree.yview)
        self.scroll_x = ttk.Scrollbar(self.frame_tree, orient="horizontal", command=self.tree.xview)

        self.tree.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.scroll_y.grid(row=0, column=1, sticky="ns")
        self.scroll_x.grid(row=1, column=0, sticky="ew")

        self.tree.bind("<Double-1>", self.abrir_arquivo_evento)

        self.frame_botoes = ctk.CTkFrame(self.frame_resultados, fg_color=COR_CARD)
        self.frame_botoes.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 12))
        self.frame_botoes.grid_columnconfigure(8, weight=1)

        self.btn_abrir_arquivo = ctk.CTkButton(
            self.frame_botoes,
            text="Abrir arquivo",
            fg_color=COR_AZUL_CLARO,
            hover_color=COR_AZUL_2,
            command=self.abrir_arquivo_selecionado,
            height=34
        )
        self.btn_abrir_arquivo.grid(row=0, column=0, padx=5, pady=6)

        self.btn_abrir_pasta = ctk.CTkButton(
            self.frame_botoes,
            text="Abrir pasta",
            fg_color=COR_AZUL_CLARO,
            hover_color=COR_AZUL_2,
            command=self.abrir_pasta_selecionada,
            height=34
        )
        self.btn_abrir_pasta.grid(row=0, column=1, padx=5, pady=6)

        self.btn_exportar = ctk.CTkButton(
            self.frame_botoes,
            text="Exportar Excel",
            fg_color=COR_SUCESSO,
            hover_color="#14532D",
            command=self.exportar_excel,
            height=34
        )
        self.btn_exportar.grid(row=0, column=2, padx=5, pady=6)

        self.btn_copiar_todos = ctk.CTkButton(
            self.frame_botoes,
            text="Copiar todos",
            fg_color=COR_ALERTA,
            hover_color="#78350F",
            command=self.copiar_todos_resultados,
            height=34
        )
        self.btn_copiar_todos.grid(row=0, column=3, padx=5, pady=6)

        self.btn_copiar_selecionados = ctk.CTkButton(
            self.frame_botoes,
            text="Copiar selecionados",
            fg_color=COR_ALERTA,
            hover_color="#78350F",
            command=self.copiar_resultados_selecionados,
            height=34
        )
        self.btn_copiar_selecionados.grid(row=0, column=4, padx=5, pady=6)

        self.btn_parar = ctk.CTkButton(
            self.frame_botoes,
            text="Parar",
            fg_color=COR_ERRO,
            hover_color="#7F1D1D",
            command=self.parar_processamento,
            height=34
        )
        self.btn_parar.grid(row=0, column=5, padx=5, pady=6)

        self.lbl_assinatura = ctk.CTkLabel(
            self.frame_botoes,
            text="Anderson Marinho | Igarapé Digital",
            text_color=COR_TEXTO_SEC,
            font=("Segoe UI", 11)
        )
        self.lbl_assinatura.grid(row=0, column=8, sticky="e", padx=6, pady=6)

        self._configurar_estilo_tabela()

    def _configurar_estilo_tabela(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Treeview",
            background="white",
            foreground=COR_TEXTO,
            rowheight=28,
            fieldbackground="white",
            bordercolor=COR_BORDA,
            borderwidth=1,
            font=("Segoe UI", 10)
        )
        style.configure(
            "Treeview.Heading",
            background=COR_AZUL,
            foreground="white",
            relief="flat",
            font=("Segoe UI", 10, "bold")
        )
        style.map(
            "Treeview.Heading",
            background=[("active", COR_AZUL_2)]
        )

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta base para busca")
        if pasta:
            self.var_pasta.set(pasta)

    def selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta destino para copiar os arquivos")
        if pasta:
            self.var_pasta_destino.set(pasta)

    def iniciar_busca(self):
        if self.busca_em_execucao:
            messagebox.showwarning("Busca em andamento", "Já existe uma busca em execução.")
            return

        if self.copia_em_execucao:
            messagebox.showwarning("Cópia em andamento", "Aguarde a cópia finalizar antes de iniciar nova busca.")
            return

        pasta = self.var_pasta.get().strip()
        termo = self.var_termo.get().strip()

        if not pasta:
            messagebox.showwarning("Pasta obrigatória", "Informe ou selecione uma pasta base.")
            return

        if not os.path.isdir(pasta):
            messagebox.showerror("Pasta inválida", "A pasta informada não existe ou não está acessível.")
            return

        if not termo:
            messagebox.showwarning("Termo obrigatório", "Informe um termo para buscar.")
            return

        if not self.var_buscar_nome.get() and not self.var_buscar_conteudo_pdf.get():
            messagebox.showwarning(
                "Opção obrigatória",
                "Marque pelo menos uma opção: buscar no nome do arquivo ou buscar dentro do PDF."
            )
            return

        if self.var_buscar_conteudo_pdf.get() and not PYPDF2_DISPONIVEL:
            messagebox.showerror(
                "PyPDF2 não instalado",
                "Para buscar dentro do PDF, instale a biblioteca com:\n\npip install PyPDF2"
            )
            return

        self.limpar_resultados()
        self.busca_em_execucao = True
        self.var_status.set("Buscando arquivos. Aguarde.")
        self.var_progresso.set(0)

        self.thread_busca = threading.Thread(target=self.executar_busca, daemon=True)
        self.thread_busca.start()

    def parar_processamento(self):
        if self.busca_em_execucao:
            self.busca_em_execucao = False
            self.var_status.set("Solicitação de parada enviada para a busca.")
            return

        if self.copia_em_execucao:
            self.copia_em_execucao = False
            self.var_status.set("Solicitação de parada enviada para a cópia.")
            return

        self.var_status.set("Nenhum processamento em execução.")

    def executar_busca(self):
        try:
            pasta_base = self.var_pasta.get().strip()
            termo = self.var_termo.get().strip()
            termo_lower = termo.lower()

            extensoes = self._normalizar_extensoes(self.var_extensoes.get())
            incluir_subpastas = self.var_incluir_subpastas.get()
            buscar_nome = self.var_buscar_nome.get()
            buscar_conteudo_pdf = self.var_buscar_conteudo_pdf.get()
            ignorar_temporarios = self.var_ignorar_temporarios.get()
            limite_resultados = self._obter_limite_resultados()

            arquivos = self._coletar_arquivos(
                pasta_base=pasta_base,
                extensoes=extensoes,
                incluir_subpastas=incluir_subpastas,
                ignorar_temporarios=ignorar_temporarios
            )

            total = len(arquivos)

            if total == 0:
                self.after(0, lambda: self.var_status.set("Nenhum arquivo encontrado para as extensões informadas."))
                self.after(0, lambda: self.var_progresso.set(0))
                self.busca_em_execucao = False
                return

            encontrados = 0
            caminhos_ja_adicionados_por_nome = set()

            for idx, caminho in enumerate(arquivos, start=1):
                if not self.busca_em_execucao:
                    break

                progresso = idx / total
                self.after(0, lambda p=progresso: self.var_progresso.set(p))
                self.after(
                    0,
                    lambda i=idx, t=total, e=encontrados: self.var_status.set(
                        f"Analisando {i} de {t}. Resultados encontrados: {e}"
                    )
                )

                if encontrados >= limite_resultados:
                    self.after(
                        0,
                        lambda: self.var_status.set(
                            f"Limite de {limite_resultados} resultados atingido. Busca finalizada."
                        )
                    )
                    break

                nome_arquivo = os.path.basename(caminho)
                ext = os.path.splitext(nome_arquivo)[1].lower()

                if buscar_nome and termo_lower in nome_arquivo.lower():
                    registro = self._montar_registro(
                        tipo="Nome do arquivo",
                        caminho=caminho,
                        pagina=""
                    )
                    self.resultados.append(registro)
                    caminhos_ja_adicionados_por_nome.add(caminho)
                    encontrados += 1
                    self.after(0, lambda r=registro: self._adicionar_linha_tabela(r))

                if encontrados >= limite_resultados:
                    continue

                if buscar_conteudo_pdf and ext == ".pdf":
                    paginas_encontradas = self._buscar_texto_em_pdf(caminho, termo_lower)

                    for pagina in paginas_encontradas:
                        if not self.busca_em_execucao:
                            break

                        if encontrados >= limite_resultados:
                            break

                        registro = self._montar_registro(
                            tipo="Conteúdo PDF",
                            caminho=caminho,
                            pagina=pagina
                        )
                        self.resultados.append(registro)
                        encontrados += 1
                        self.after(0, lambda r=registro: self._adicionar_linha_tabela(r))

            self.busca_em_execucao = False
            self.after(0, lambda: self.var_progresso.set(1))

            if len(self.resultados) == 0:
                self.after(0, lambda: self.var_status.set("Busca concluída. Nenhum resultado encontrado."))
            else:
                self.after(
                    0,
                    lambda: self.var_status.set(
                        f"Busca concluída. {len(self.resultados)} resultado(s) encontrado(s)."
                    )
                )

        except Exception as e:
            self.busca_em_execucao = False
            erro = f"Erro durante a busca: {str(e)}"
            print(traceback.format_exc())
            self.after(0, lambda: self.var_status.set(erro))
            self.after(0, lambda: messagebox.showerror("Erro", erro))

    def _normalizar_extensoes(self, texto_extensoes):
        texto = texto_extensoes.strip().lower()

        if not texto:
            return []

        partes = []
        for item in texto.replace(",", ";").split(";"):
            item = item.strip().lower()
            if not item:
                continue
            if not item.startswith("."):
                item = "." + item
            partes.append(item)

        return list(dict.fromkeys(partes))

    def _obter_limite_resultados(self):
        try:
            limite = int(self.var_limite_resultados.get().strip())
            if limite <= 0:
                return 1000
            return limite
        except Exception:
            return 1000

    def _coletar_arquivos(self, pasta_base, extensoes, incluir_subpastas, ignorar_temporarios):
        arquivos_encontrados = []

        if incluir_subpastas:
            for raiz, dirs, files in os.walk(pasta_base):
                if not self.busca_em_execucao:
                    break

                dirs[:] = [
                    d for d in dirs
                    if not self._deve_ignorar_pasta(d)
                ]

                for nome in files:
                    if ignorar_temporarios and self._deve_ignorar_arquivo(nome):
                        continue

                    caminho = os.path.join(raiz, nome)

                    if self._arquivo_tem_extensao(caminho, extensoes):
                        arquivos_encontrados.append(caminho)
        else:
            for nome in os.listdir(pasta_base):
                caminho = os.path.join(pasta_base, nome)
                if os.path.isfile(caminho):
                    if ignorar_temporarios and self._deve_ignorar_arquivo(nome):
                        continue
                    if self._arquivo_tem_extensao(caminho, extensoes):
                        arquivos_encontrados.append(caminho)

        return arquivos_encontrados

    def _deve_ignorar_pasta(self, nome_pasta):
        nome = nome_pasta.lower().strip()

        pastas_ignoradas = {
            ".git",
            "__pycache__",
            "node_modules",
            ".venv",
            "venv",
            "env",
            "$recycle.bin",
            "system volume information"
        }

        if nome in pastas_ignoradas:
            return True

        if nome.startswith("~"):
            return True

        return False

    def _deve_ignorar_arquivo(self, nome_arquivo):
        nome = nome_arquivo.lower().strip()

        if nome.startswith("~$"):
            return True

        if nome.endswith(".tmp"):
            return True

        if nome.endswith(".crdownload"):
            return True

        if nome.endswith(".partial"):
            return True

        if nome == "desktop.ini":
            return True

        return False

    def _arquivo_tem_extensao(self, caminho, extensoes):
        if not extensoes:
            return True

        ext = os.path.splitext(caminho)[1].lower()
        return ext in extensoes

    def _buscar_texto_em_pdf(self, caminho_pdf, termo_lower):
        paginas = []

        try:
            with open(caminho_pdf, "rb") as arquivo:
                leitor = PyPDF2.PdfReader(arquivo)

                for indice, pagina in enumerate(leitor.pages, start=1):
                    if not self.busca_em_execucao:
                        break

                    try:
                        texto = pagina.extract_text() or ""
                    except Exception:
                        texto = ""

                    if termo_lower in texto.lower():
                        paginas.append(str(indice))

        except Exception:
            pass

        return paginas

    def _montar_registro(self, tipo, caminho, pagina):
        try:
            stat = os.stat(caminho)
            tamanho_kb = round(stat.st_size / 1024, 2)
            modificado = datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M")
        except Exception:
            tamanho_kb = ""
            modificado = ""

        nome_arquivo = os.path.basename(caminho)
        pasta = os.path.dirname(caminho)
        extensao = os.path.splitext(nome_arquivo)[1].lower()

        return {
            "Tipo": tipo,
            "Arquivo": nome_arquivo,
            "Extensao": extensao,
            "Pasta": pasta,
            "Pagina": pagina,
            "Tamanho_KB": tamanho_kb,
            "Modificado": modificado,
            "Caminho": caminho
        }

    def _adicionar_linha_tabela(self, registro):
        self.tree.insert(
            "",
            "end",
            values=(
                registro["Tipo"],
                registro["Arquivo"],
                registro["Extensao"],
                registro["Pasta"],
                registro["Pagina"],
                registro["Tamanho_KB"],
                registro["Modificado"],
                registro["Caminho"]
            )
        )

    def limpar_resultados(self):
        if self.busca_em_execucao or self.copia_em_execucao:
            messagebox.showwarning("Processamento em execução", "Pare ou aguarde o processamento antes de limpar.")
            return

        self.resultados = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.var_progresso.set(0)
        self.var_status.set("Resultados limpos. Pronto para buscar.")

    def obter_registro_selecionado(self):
        selecionado = self.tree.selection()

        if not selecionado:
            messagebox.showwarning("Seleção obrigatória", "Selecione um resultado na tabela.")
            return None

        valores = self.tree.item(selecionado[0], "values")

        if not valores:
            return None

        caminho = valores[7]

        for registro in self.resultados:
            if registro["Caminho"] == caminho:
                return registro

        return {
            "Caminho": caminho,
            "Pasta": valores[3],
            "Arquivo": valores[1]
        }

    def obter_registros_selecionados(self):
        selecionados = self.tree.selection()

        if not selecionados:
            return []

        caminhos = []

        for item in selecionados:
            valores = self.tree.item(item, "values")
            if valores:
                caminhos.append(valores[7])

        registros = []
        caminhos_adicionados = set()

        for caminho in caminhos:
            if caminho in caminhos_adicionados:
                continue

            for registro in self.resultados:
                if registro["Caminho"] == caminho:
                    registros.append(registro)
                    caminhos_adicionados.add(caminho)
                    break

        return registros

    def abrir_arquivo_evento(self, event):
        self.abrir_arquivo_selecionado()

    def abrir_arquivo_selecionado(self):
        registro = self.obter_registro_selecionado()
        if not registro:
            return

        caminho = registro["Caminho"]

        if not os.path.isfile(caminho):
            messagebox.showerror("Arquivo não encontrado", "O arquivo selecionado não existe mais no caminho informado.")
            return

        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro ao abrir", f"Não foi possível abrir o arquivo.\n\n{str(e)}")

    def abrir_pasta_selecionada(self):
        registro = self.obter_registro_selecionado()
        if not registro:
            return

        caminho = registro["Caminho"]
        pasta = os.path.dirname(caminho)

        if not os.path.isdir(pasta):
            messagebox.showerror("Pasta não encontrada", "A pasta do arquivo selecionado não existe mais.")
            return

        try:
            os.startfile(pasta)
        except Exception as e:
            messagebox.showerror("Erro ao abrir pasta", f"Não foi possível abrir a pasta.\n\n{str(e)}")

    def exportar_excel(self):
        if not self.resultados:
            messagebox.showwarning("Sem resultados", "Não há resultados para exportar.")
            return

        caminho_saida = filedialog.asksaveasfilename(
            title="Salvar resultado da busca",
            defaultextension=".xlsx",
            filetypes=[("Arquivo Excel", "*.xlsx")],
            initialfile=f"resultado_busca_onedrive_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if not caminho_saida:
            return

        try:
            df = pd.DataFrame(self.resultados)

            with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="RESULTADOS")
                ws = writer.book["RESULTADOS"]
                self._formatar_planilha(ws)

            messagebox.showinfo("Exportado", f"Resultado exportado com sucesso:\n\n{caminho_saida}")

        except Exception as e:
            messagebox.showerror("Erro ao exportar", f"Não foi possível exportar o Excel.\n\n{str(e)}")

    def copiar_todos_resultados(self):
        if not self.resultados:
            messagebox.showwarning("Sem resultados", "Não há arquivos encontrados para copiar.")
            return

        registros_unicos = self._obter_registros_unicos_por_caminho(self.resultados)
        self._iniciar_copia(registros_unicos, tipo_copia="todos")

    def copiar_resultados_selecionados(self):
        registros = self.obter_registros_selecionados()

        if not registros:
            messagebox.showwarning("Seleção obrigatória", "Selecione um ou mais arquivos na tabela para copiar.")
            return

        registros_unicos = self._obter_registros_unicos_por_caminho(registros)
        self._iniciar_copia(registros_unicos, tipo_copia="selecionados")

    def _iniciar_copia(self, registros, tipo_copia):
        if self.busca_em_execucao:
            messagebox.showwarning("Busca em andamento", "Aguarde a busca finalizar antes de copiar.")
            return

        if self.copia_em_execucao:
            messagebox.showwarning("Cópia em andamento", "Já existe uma cópia em execução.")
            return

        pasta_destino = self.var_pasta_destino.get().strip()

        if not pasta_destino:
            pasta_destino = filedialog.askdirectory(title="Selecione a pasta destino para copiar os arquivos")
            if not pasta_destino:
                return
            self.var_pasta_destino.set(pasta_destino)

        try:
            os.makedirs(pasta_destino, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Pasta destino inválida", f"Não foi possível criar ou acessar a pasta destino.\n\n{str(e)}")
            return

        if not os.path.isdir(pasta_destino):
            messagebox.showerror("Pasta destino inválida", "A pasta destino informada não existe ou não está acessível.")
            return

        pergunta = messagebox.askyesno(
            "Confirmar cópia",
            f"Deseja copiar {len(registros)} arquivo(s) para:\n\n{pasta_destino}?"
        )

        if not pergunta:
            return

        self.copia_em_execucao = True
        self.var_progresso.set(0)
        self.var_status.set("Copiando arquivos. Aguarde.")

        self.thread_copia = threading.Thread(
            target=self._executar_copia,
            args=(registros, pasta_destino, tipo_copia),
            daemon=True
        )
        self.thread_copia.start()

    def _executar_copia(self, registros, pasta_destino, tipo_copia):
        log_copia = []

        try:
            total = len(registros)

            for idx, registro in enumerate(registros, start=1):
                if not self.copia_em_execucao:
                    break

                origem = registro["Caminho"]
                nome_arquivo = os.path.basename(origem)

                progresso = idx / total
                self.after(0, lambda p=progresso: self.var_progresso.set(p))
                self.after(
                    0,
                    lambda i=idx, t=total, n=nome_arquivo: self.var_status.set(
                        f"Copiando {i} de {t}: {n}"
                    )
                )

                item_log = {
                    "DataHora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    "TipoCopia": tipo_copia,
                    "Arquivo": nome_arquivo,
                    "Origem": origem,
                    "Destino": "",
                    "Status": "",
                    "Mensagem": ""
                }

                try:
                    if not os.path.isfile(origem):
                        item_log["Status"] = "Erro"
                        item_log["Mensagem"] = "Arquivo de origem não encontrado."
                        log_copia.append(item_log)
                        continue

                    destino_final = self._gerar_caminho_destino_sem_sobrescrever(
                        pasta_destino=pasta_destino,
                        nome_arquivo=nome_arquivo
                    )

                    shutil.copy2(origem, destino_final)

                    item_log["Destino"] = destino_final
                    item_log["Status"] = "Copiado"
                    item_log["Mensagem"] = "Arquivo copiado com sucesso."
                    log_copia.append(item_log)

                except Exception as e:
                    item_log["Status"] = "Erro"
                    item_log["Mensagem"] = str(e)
                    log_copia.append(item_log)

            self.copia_em_execucao = False
            self.after(0, lambda: self.var_progresso.set(1))

            copiados = len([x for x in log_copia if x["Status"] == "Copiado"])
            erros = len([x for x in log_copia if x["Status"] == "Erro"])

            caminho_log = self._salvar_log_copia(pasta_destino, log_copia)

            mensagem_final = (
                f"Cópia concluída. Copiados: {copiados}. Erros: {erros}.\n"
                f"Log salvo em: {caminho_log}"
            )

            self.after(0, lambda: self.var_status.set(mensagem_final))
            self.after(0, lambda: messagebox.showinfo("Cópia concluída", mensagem_final))

        except Exception as e:
            self.copia_em_execucao = False
            erro = f"Erro durante a cópia: {str(e)}"
            print(traceback.format_exc())
            self.after(0, lambda: self.var_status.set(erro))
            self.after(0, lambda: messagebox.showerror("Erro", erro))

    def _obter_registros_unicos_por_caminho(self, registros):
        unicos = []
        vistos = set()

        for registro in registros:
            caminho = registro.get("Caminho", "")
            if not caminho:
                continue

            if caminho in vistos:
                continue

            vistos.add(caminho)
            unicos.append(registro)

        return unicos

    def _gerar_caminho_destino_sem_sobrescrever(self, pasta_destino, nome_arquivo):
        nome_base, extensao = os.path.splitext(nome_arquivo)
        destino = os.path.join(pasta_destino, nome_arquivo)

        if not os.path.exists(destino):
            return destino

        contador = 1

        while True:
            novo_nome = f"{nome_base}_{contador:03d}{extensao}"
            novo_destino = os.path.join(pasta_destino, novo_nome)

            if not os.path.exists(novo_destino):
                return novo_destino

            contador += 1

    def _salvar_log_copia(self, pasta_destino, log_copia):
        caminho_log = os.path.join(
            pasta_destino,
            f"log_copia_buscador_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        try:
            df = pd.DataFrame(log_copia)

            with pd.ExcelWriter(caminho_log, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="LOG_COPIA")
                ws = writer.book["LOG_COPIA"]
                self._formatar_planilha(ws)

            return caminho_log

        except Exception:
            return "Não foi possível gerar o log em Excel."

    def _formatar_planilha(self, ws):
        for coluna in ws.columns:
            maior = 0
            letra = coluna[0].column_letter

            for celula in coluna:
                try:
                    valor = str(celula.value) if celula.value is not None else ""
                    if len(valor) > maior:
                        maior = len(valor)
                except Exception:
                    pass

            largura = min(max(maior + 2, 12), 90)
            ws.column_dimensions[letra].width = largura

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions


def main():
    app = BuscadorOneDrivePDF()
    app.mainloop()


if __name__ == "__main__":
    main()

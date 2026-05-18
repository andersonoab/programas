import os
import customtkinter as ctk
from tkinter import filedialog, messagebox


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class CustomerThinkerAFDRemover(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CustomerThinker | Removedor de CRC AFD")
        self.geometry("920x700")
        self.minsize(920, 700)
        self.configure(fg_color="#F4F7FB")

        self.arquivo_entrada = ""
        self.arquivo_saida = ""

        self._montar_interface()

    def _montar_interface(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        topo = ctk.CTkFrame(self, fg_color="#121C4E", corner_radius=0, height=82)
        topo.grid(row=0, column=0, sticky="ew")
        topo.grid_columnconfigure(0, weight=1)

        titulo = ctk.CTkLabel(
            topo,
            text="CustomerThinker | Tratamento de AFD",
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color="white"
        )
        titulo.grid(row=0, column=0, padx=20, pady=(14, 2), sticky="w")

        subtitulo = ctk.CTkLabel(
            topo,
            text="Remoção dos 4 últimos caracteres com validação estrutural e contagem real do arquivo",
            font=ctk.CTkFont(size=14),
            text_color="#D9E2FF"
        )
        subtitulo.grid(row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        corpo = ctk.CTkFrame(self, fg_color="transparent")
        corpo.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        corpo.grid_columnconfigure(0, weight=1)
        corpo.grid_rowconfigure(3, weight=1)

        card_entrada = ctk.CTkFrame(
            corpo, fg_color="white", corner_radius=16,
            border_width=1, border_color="#D9E2EC"
        )
        card_entrada.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        card_entrada.grid_columnconfigure(0, weight=1)

        lbl_entrada = ctk.CTkLabel(
            card_entrada,
            text="Arquivo de entrada",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1F2937"
        )
        lbl_entrada.grid(row=0, column=0, padx=20, pady=(18, 8), sticky="w")

        self.entry_entrada = ctk.CTkEntry(
            card_entrada,
            placeholder_text="Selecione o arquivo AFD .txt",
            height=40,
            font=ctk.CTkFont(size=14),
            border_color="#B8C4D6"
        )
        self.entry_entrada.grid(row=1, column=0, padx=(20, 220), pady=(0, 18), sticky="ew")

        btn_entrada = ctk.CTkButton(
            card_entrada,
            text="Selecionar arquivo",
            width=180,
            height=40,
            fg_color="#1E3A8A",
            hover_color="#163172",
            command=self.selecionar_arquivo_entrada
        )
        btn_entrada.grid(row=1, column=0, padx=(0, 20), pady=(0, 18), sticky="e")

        card_saida = ctk.CTkFrame(
            corpo, fg_color="white", corner_radius=16,
            border_width=1, border_color="#D9E2EC"
        )
        card_saida.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        card_saida.grid_columnconfigure(0, weight=1)

        lbl_saida = ctk.CTkLabel(
            card_saida,
            text="Arquivo de saída",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1F2937"
        )
        lbl_saida.grid(row=0, column=0, padx=20, pady=(18, 8), sticky="w")

        self.entry_saida = ctk.CTkEntry(
            card_saida,
            placeholder_text="Escolha onde salvar o arquivo tratado",
            height=40,
            font=ctk.CTkFont(size=14),
            border_color="#B8C4D6"
        )
        self.entry_saida.grid(row=1, column=0, padx=(20, 220), pady=(0, 18), sticky="ew")

        btn_saida = ctk.CTkButton(
            card_saida,
            text="Escolher destino",
            width=180,
            height=40,
            fg_color="#1E3A8A",
            hover_color="#163172",
            command=self.selecionar_arquivo_saida
        )
        btn_saida.grid(row=1, column=0, padx=(0, 20), pady=(0, 18), sticky="e")

        card_acoes = ctk.CTkFrame(
            corpo, fg_color="white", corner_radius=16,
            border_width=1, border_color="#D9E2EC"
        )
        card_acoes.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        card_acoes.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_processar = ctk.CTkButton(
            card_acoes,
            text="Remover 4 dígitos finais",
            height=45,
            fg_color="#0F766E",
            hover_color="#0B5E58",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.processar_arquivo
        )
        self.btn_processar.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        btn_limpar = ctk.CTkButton(
            card_acoes,
            text="Limpar campos",
            height=45,
            fg_color="#64748B",
            hover_color="#475569",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.limpar_campos
        )
        btn_limpar.grid(row=0, column=1, padx=(0, 20), pady=20, sticky="ew")

        btn_sugerir = ctk.CTkButton(
            card_acoes,
            text="Sugerir nome de saída",
            height=45,
            fg_color="#1D4ED8",
            hover_color="#1E40AF",
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.sugerir_saida
        )
        btn_sugerir.grid(row=0, column=2, padx=(0, 20), pady=20, sticky="ew")

        card_log = ctk.CTkFrame(
            corpo, fg_color="white", corner_radius=16,
            border_width=1, border_color="#D9E2EC"
        )
        card_log.grid(row=3, column=0, sticky="nsew")
        card_log.grid_columnconfigure(0, weight=1)
        card_log.grid_rowconfigure(1, weight=1)

        lbl_log = ctk.CTkLabel(
            card_log,
            text="Log de execução",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#1F2937"
        )
        lbl_log.grid(row=0, column=0, padx=20, pady=(18, 10), sticky="w")

        self.txt_log = ctk.CTkTextbox(
            card_log,
            height=320,
            font=ctk.CTkFont(size=13),
            border_width=1,
            border_color="#D9E2EC"
        )
        self.txt_log.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")

        self.log("Sistema iniciado.")
        self.log("Selecione um arquivo AFD e informe o destino do arquivo tratado.")
        self.log("Contagem principal: total de linhas reais do arquivo.")
        self.log("Header = primeira linha.")
        self.log("Trailer = linha gerada automaticamente com a contagem (excluindo header).")

        rodape = ctk.CTkFrame(self, fg_color="#EAF0FA", corner_radius=0, height=34)
        rodape.grid(row=2, column=0, sticky="ew")
        rodape.grid_columnconfigure(0, weight=1)

        lbl_rodape = ctk.CTkLabel(
            rodape,
            text="Anderson Marinho | Igarapé Digital",
            font=ctk.CTkFont(size=12),
            text_color="#334155"
        )
        lbl_rodape.grid(row=0, column=0, padx=15, pady=7, sticky="e")

    def log(self, mensagem):
        self.txt_log.insert("end", mensagem + "\n")
        self.txt_log.see("end")
        self.update_idletasks()

    def selecionar_arquivo_entrada(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo AFD",
            filetypes=[("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")]
        )

        if caminho:
            self.arquivo_entrada = caminho
            self.entry_entrada.delete(0, "end")
            self.entry_entrada.insert(0, caminho)
            self.log(f"Arquivo de entrada selecionado: {caminho}")

            if not self.entry_saida.get().strip():
                self.sugerir_saida()

    def selecionar_arquivo_saida(self):
        caminho = filedialog.asksaveasfilename(
            title="Salvar arquivo tratado",
            defaultextension=".txt",
            filetypes=[("Arquivos TXT", "*.txt"), ("Todos os arquivos", "*.*")],
            initialfile="afd_tratado.txt"
        )

        if caminho:
            self.arquivo_saida = caminho
            self.entry_saida.delete(0, "end")
            self.entry_saida.insert(0, caminho)
            self.log(f"Arquivo de saída definido: {caminho}")

    def sugerir_saida(self):
        entrada = self.entry_entrada.get().strip()

        if not entrada:
            self.log("Nenhum arquivo de entrada selecionado para sugerir saída.")
            return

        pasta = os.path.dirname(entrada)
        nome = os.path.basename(entrada)
        nome_sem_ext, ext = os.path.splitext(nome)

        sugestao = os.path.join(pasta, f"{nome_sem_ext}_sem_crc{ext}")

        self.arquivo_saida = sugestao
        self.entry_saida.delete(0, "end")
        self.entry_saida.insert(0, sugestao)
        self.log(f"Nome de saída sugerido: {sugestao}")

    def limpar_campos(self):
        self.arquivo_entrada = ""
        self.arquivo_saida = ""
        self.entry_entrada.delete(0, "end")
        self.entry_saida.delete(0, "end")
        self.txt_log.delete("1.0", "end")
        self.log("Campos limpos.")
        self.log("Pronto para nova execução.")

    def gerar_linha_trailer(self, contagem):
        """
        Gera a linha de trailer no formato AFD:
        999999999 + 000000000 + contagem(9 dígitos) + 000000000 + 000000000
        Total: 45 caracteres (5 blocos de 9).
        """
        bloco1 = "999999999"
        bloco2 = "000000000"
        bloco3 = f"{contagem:09d}"
        bloco4 = "000000000"
        bloco5 = "000000000"
        return bloco1 + bloco2 + bloco3 + bloco4 + bloco5

    def analisar_estrutura_afd(self, linhas):
        total_linhas = len(linhas)

        possui_header = total_linhas >= 1

        header = linhas[0].rstrip("\n").rstrip("\r") if possui_header else ""

        # Contagem de linhas excluindo o header (primeira linha)
        linhas_sem_header = max(total_linhas - 1, 0)

        return {
            "total_linhas": total_linhas,
            "possui_header": possui_header,
            "header": header,
            "linhas_sem_header": linhas_sem_header,
        }

    def remover_crc_afd(self, arquivo_entrada, arquivo_saida):
        with open(arquivo_entrada, "r", encoding="latin-1") as f:
            linhas = f.readlines()

        analise = self.analisar_estrutura_afd(linhas)

        novas_linhas = []
        linhas_tratadas = 0

        for linha in linhas:
            linha_limpa = linha.rstrip("\n").rstrip("\r")

            if len(linha_limpa) > 4:
                linha_tratada = linha_limpa[:-4]
                linhas_tratadas += 1
            else:
                linha_tratada = linha_limpa

            novas_linhas.append(linha_tratada + "\n")

        # Contagem: todas as linhas excluindo a primeira (header)
        contagem_trailer = len(novas_linhas) - 1 if len(novas_linhas) > 1 else 0

        # Gera e adiciona a linha de trailer ao final
        linha_trailer = self.gerar_linha_trailer(contagem_trailer)
        novas_linhas.append(linha_trailer + "\n")

        with open(arquivo_saida, "w", encoding="latin-1") as f:
            f.writelines(novas_linhas)

        analise["contagem_trailer"] = contagem_trailer
        analise["linha_trailer"] = linha_trailer

        return analise, linhas_tratadas

    def processar_arquivo(self):
        arquivo_entrada = self.entry_entrada.get().strip()
        arquivo_saida = self.entry_saida.get().strip()

        if not arquivo_entrada:
            messagebox.showwarning("Atenção", "Selecione o arquivo de entrada.")
            self.log("Processo interrompido: arquivo de entrada não informado.")
            return

        if not os.path.exists(arquivo_entrada):
            messagebox.showerror("Erro", "O arquivo de entrada não existe.")
            self.log("Erro: o arquivo de entrada informado não existe.")
            return

        if not arquivo_saida:
            messagebox.showwarning("Atenção", "Informe o arquivo de saída.")
            self.log("Processo interrompido: arquivo de saída não informado.")
            return

        try:
            pasta_saida = os.path.dirname(arquivo_saida)
            if pasta_saida and not os.path.exists(pasta_saida):
                os.makedirs(pasta_saida, exist_ok=True)
                self.log(f"Pasta de saída criada: {pasta_saida}")

            self.btn_processar.configure(state="disabled", text="Processando...")
            self.log("Iniciando tratamento do arquivo...")

            analise, linhas_tratadas = self.remover_crc_afd(arquivo_entrada, arquivo_saida)

            self.log("")
            self.log("═══════════════════════════════════════════")
            self.log("  Resumo estrutural do AFD")
            self.log("═══════════════════════════════════════════")
            self.log(f"  Total de linhas do arquivo original: {analise['total_linhas']}")
            self.log(f"  Linhas sem o header (primeira linha): {analise['linhas_sem_header']}")
            self.log(f"  Linhas tratadas (remoção dos 4 últimos): {linhas_tratadas}")
            self.log("───────────────────────────────────────────")
            self.log(f"  Trailer gerado automaticamente:")
            self.log(f"  Contagem no trailer: {analise['contagem_trailer']}")
            self.log(f"  Linha trailer: {analise['linha_trailer']}")
            self.log("═══════════════════════════════════════════")
            self.log(f"Arquivo gerado com sucesso: {arquivo_saida}")

            self.btn_processar.configure(state="normal", text="Remover 4 dígitos finais")

            messagebox.showinfo(
                "Sucesso",
                f"Arquivo gerado com sucesso.\n\n{arquivo_saida}\n\n"
                f"Total de linhas original: {analise['total_linhas']}\n"
                f"Linhas sem header: {analise['linhas_sem_header']}\n"
                f"Contagem no trailer: {analise['contagem_trailer']}\n"
                f"Trailer: {analise['linha_trailer']}"
            )

        except Exception as e:
            self.btn_processar.configure(state="normal", text="Remover 4 dígitos finais")
            self.log(f"Erro durante o processamento: {str(e)}")
            messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo.\n\n{str(e)}")


if __name__ == "__main__":
    app = CustomerThinkerAFDRemover()
    app.mainloop()

'''
& C:/Users/nome/AppData/Local/Programs/Python/Python310/python.exe -m cx_Freeze `
--script "c:/_RPA\aAppEnvios PDF/AFD.py" `
--target-dir dist_afd `
--target-name AFD_App.exe `
--base-name Win32GUI

'''

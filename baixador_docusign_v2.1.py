# -*- coding: utf-8 -*-
"""
Baixador DocuSign — Envelopes Concluídos (Sonova)  v3.5.2
Automação via Playwright usando o Edge instalado (channel="msedge"), sem API.
UI CustomTkinter no padrão Sonova OFC.

v2: seletores reais do DocuSign via data-qa
    (manage-envelopes-list-row-<ENVELOPE_ID>-actions-download_envelope-text),
    chave de deduplicação = Envelope ID, botão Diagnóstico, sem aviso --no-sandbox,
    barra de botões em duas linhas.
v3.3: base = v2 (leitura da lista, filtros, seleção, histórico e CSV inalterados).
    + PDF unificado por envelope: usa os seletores reais do modal
        checkbox -> input#download-combined-label
        rótulo   -> [data-qa="download-combined-label-label"]
        botão    -> [data-qa="download-document-button"]
      Lê o estado real do checkbox, ajusta pelo rótulo só se precisar e CONFIRMA
      antes de clicar em Baixar. Se não conseguir confirmar, não baixa aquele
      envelope (marca Erro) para não gerar arquivo no formato errado.
    + Parar: cancela a espera pelo arquivo e o download que chegar depois do
      pedido; arquivos já concluídos permanecem na pasta e no log.
v3.4: baseada na versão v3.3 corrigida e validada pelo Anderson.
    + Nome: data, título e colaborador, sem textos de seleção da lista;
      remove caracteres inválidos e considera o caminho completo da pasta.
    + Salvamento: não sobrescreve arquivos; publica a cópia completa antes
      de registrar o histórico. Falhas preservam a cópia temporária e mostram
      o caminho para recuperação. Leitura, Edge e seletores preservados.
    + Parar: espera cancelável também durante a transferência; save_as só
      recebe o arquivo depois da conclusão do download nativo do Edge.
v3.5: corrige o WinError 5 no salvamento relatado no Windows.
    + Grava diretamente no nome final com criação exclusiva, sem arquivo
      vazio de reserva seguido de os.replace. Nunca sobrescreve outro arquivo.
    + Bloqueios de acesso têm novas tentativas usando o mesmo PDF local;
      Parar também interrompe a cópia e a espera entre essas tentativas.
    + Só confirma Baixado depois de fechar e conferir a cópia completa.
v3.5.1: única correção funcional = mantém pageSize=50 na URL da lista,
    inclusive ao carregar uma configuração antiga ou editar o campo URL.
v3.5.2: escolha explícita do formato pelo checkbox na tela do programa.
    Marcado = PDF único por envelope; desmarcado = documentos sem combinar.
    Salva a escolha ao clicar e confirma o mesmo estado no modal do DocuSign.

Dependências:  python.exe -m pip install customtkinter playwright
(NÃO é necessário rodar "playwright install": usa o Edge da máquina.)

Autor: Anderson Souza
"""

import csv
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import customtkinter as ctk

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    PLAYWRIGHT_OK = True
except ImportError:
    PLAYWRIGHT_OK = False

# ════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO PADRÃO
# ════════════════════════════════════════════════════════════════════
APP_NOME = "Baixador DocuSign — Envelopes Concluídos"
ARQ_CONFIG = Path.home() / ".docusign_baixador_config.json"
PERFIL_EDGE = Path.home() / "PlaywrightEdgeDocusign"
PASTA_TEMP = Path.home() / "Downloads" / "_docusign_temp"

CONFIG_PADRAO = {
    "url": "https://apps.docusign.com/send/documents?status=completed&pageSize=50",
    "destino": r"C:\Users\99andsouza\Sonova\BR-SO-HR-RecursosHumanos - General\ADMISSÕES\2026\Docusign Concluido",
    "filtro": r"Admiss",
    "intervalo": 3,
    "combinar": True,
    "ignorar_baixados": True,
}

# Seletores reais do DocuSign (data-qa)
PREFIXO_ROW = "manage-envelopes-list-row-"
SUFIXO_DL = "-actions-download_envelope-text"
SEL_BAIXAR_LISTA = f'[data-qa^="{PREFIXO_ROW}"][data-qa$="{SUFIXO_DL}"]'

# Modal "Fazer download dos arquivos" (seletores reais enviados pelo Anderson)
SEL_COMB_INPUT = "#download-combined-label"
SEL_COMB_LABEL = '[data-qa="download-combined-label-label"]'
SEL_MODAL_BTN = '[data-qa="download-document-button"]'
ESPERA_DOWNLOAD_S = 120   # mesmo limite da v2 (120000 ms)
ESPERA_TRANSFERENCIA_S = 600  # arquivos grandes: até 10 min, com Parar disponível
ESPERAS_SALVAMENTO = (0.3, 0.7, 1.5, 3.0)  # até cinco tentativas

# JS executado no navegador: para cada "Baixar" da lista, extrai o Envelope ID
# do data-qa e sobe no DOM até achar o bloco da linha (que contém a data).
JS_LER_LINHAS = r"""
els => els.map(el => {
  const qa = el.getAttribute('data-qa') || '';
  const m = qa.match(/manage-envelopes-list-row-(.+)-actions-download_envelope-text/);
  let n = el, txt = '';
  for (let i = 0; i < 20 && n.parentElement; i++) {
    n = n.parentElement;
    const t = (n.innerText || '');
    const linhas = t.split('\n').filter(x => x.trim());
    if (/\d{2}\/\d{2}\/\d{4}/.test(t) && linhas.length >= 3) { txt = t; break; }
  }
  return { guid: m ? m[1] : '', texto: txt };
})
"""

# Paleta Sonova OFC
AZUL = "#0083CA"
AZUL_ESC = "#003C64"
PETROLEO = "#005A64"
AZUL_CLARO = "#6EB4DC"
VINHO = "#7D0041"
TERRA = "#8C321E"
TXT = "#333333"
TXT2 = "#646464"
BORDA = "#CCCCCC"
BRANCO = "#FFFFFF"
VERDE = "#2E8B57"
VERDE_HOVER = "#246E45"
CINZA = "#8A8F94"
CINZA_HOVER = "#6F7479"
VERMELHO = "#C0392B"
VERMELHO_HOVER = "#992D22"
TH_BG = "#E8F3FA"
ZEBRA = "#F7F9FB"
TRACK = "#EEF3F7"
THUMB = "#B9CBD8"

FONTE_TIT = "Rufina"
FONTE_CORPO = "Roboto Light"

RE_DATA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
RE_HORA = re.compile(r"^\d{2}:\d{2}$")
RUIDO = {"concluído", "concluido", "baixar", "completed", "download"}
RUIDO.update({"desmarcar linha", "selecionar linha", "marcar linha", "selecionar",
              "mais opções", "mais opcoes", "deselect row", "select row"})
# Margem para instalações Windows que não habilitaram caminhos longos.
# O limite inclui a pasta; caracteres fora do BMP ocupam duas unidades UTF-16.
LIMITE_CAMINHO = 240
NOMES_RESERVADOS = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"}
NOMES_RESERVADOS.update(f"{p}{n}" for p in ("COM", "LPT") for n in "123456789¹²³")


# ════════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ════════════════════════════════════════════════════════════════════
def url_com_50_envelopes(url):
    url = str(url or "").strip()
    try:
        partes = urlsplit(url)
    except ValueError:
        return url
    host = (partes.hostname or "").lower()
    docusign = any(host == dominio or host.endswith("." + dominio)
                   for dominio in ("docusign.com", "docusign.net", "docusign.com.br"))
    if not docusign or partes.path.rstrip("/") != "/send/documents":
        return url
    parametros = [(k, v) for k, v in parse_qsl(partes.query, keep_blank_values=True)
                  if k.lower() != "pagesize"]
    parametros.append(("pageSize", "50"))
    return urlunsplit(partes._replace(query=urlencode(parametros)))


def carregar_config():
    cfg = dict(CONFIG_PADRAO)
    if ARQ_CONFIG.exists():
        try:
            cfg.update(json.loads(ARQ_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    cfg["url"] = url_com_50_envelopes(cfg["url"])
    return cfg


def salvar_config(cfg):
    try:
        ARQ_CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def texto_modo_download(combinar):
    if combinar:
        return "Modo escolhido: PDF único por envelope."
    return "Modo escolhido: documentos separados (o DocuSign pode enviar um ZIP)."


def limpar_nome(txt):
    txt = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', " ", str(txt or ""))
    txt = re.sub(r"\s+", " ", txt).strip(" .")
    if txt.split(".", 1)[0].upper() in NOMES_RESERVADOS:
        txt = "_" + txt
    return txt


def _unidades(txt):
    return len(txt.encode("utf-16-le")) // 2


def _cortar_nome(txt, limite):
    return txt.encode("utf-16-le")[:max(0, limite) * 2].decode("utf-16-le", errors="ignore").rstrip(" .")


def _nome_no_limite(pasta, nome, sufixo=""):
    """Encurta o título antes do colaborador, sem alterar a extensão."""
    nome = limpar_nome(nome)
    ext = Path(nome).suffix
    base = nome[:-len(ext)] if ext else nome
    limite = min(255, LIMITE_CAMINHO - _unidades(str(Path(pasta).absolute())) - 1)
    disponivel = limite - _unidades(ext + sufixo)
    if disponivel < 16:
        raise OSError("O caminho da pasta é longo demais. Escolha uma pasta com caminho mais curto.")
    if _unidades(base) > disponivel:
        if " - " in base:
            titulo, colab = base.rsplit(" - ", 1)
            # Reserva espaço para a data e parte legível do título.
            colab = _cortar_nome(colab, max(0, disponivel - 30))
            complemento = f" - {colab}" if colab else ""
            base = _cortar_nome(titulo, disponivel - _unidades(complemento)) + complemento
        else:
            base = _cortar_nome(base, disponivel)
    return f"{base}{sufixo}{ext}"


def montar_nome_arquivo(item, sugerido):
    sugerido = Path(str(sugerido or "").replace("\\", "/")).name
    ext = Path(sugerido).suffix or ".pdf"
    titulo = item.get("nome", "").strip()
    if not titulo or titulo.casefold() in RUIDO:
        titulo = Path(sugerido).stem or f"Envelope {item.get('guid') or 'sem-titulo'}"
    titulo = limpar_nome(titulo)
    if titulo.lower().endswith(ext.lower()):
        titulo = titulo[:-len(ext)].rstrip(" .")
    colab = limpar_nome(re.sub(r"\s*\+\d+\s*mais\s*$", "", item.get("colaborador") or "", flags=re.I))
    if colab and colab.casefold() not in titulo.casefold():
        titulo += f" - {colab}"
    try:
        data_iso = datetime.strptime(item.get("data") or "", "%d/%m/%Y").strftime("%Y-%m-%d")
    except ValueError:
        data_iso = "sem-data"
    return f"{data_iso}_{titulo or 'Envelope'}{ext}"


def caminho_unico(pasta: Path, nome: str) -> Path:
    pasta = Path(pasta)
    alvo = pasta / _nome_no_limite(pasta, nome)
    if not alvo.exists():
        return alvo
    n = 2
    while (pasta / _nome_no_limite(pasta, nome, f"_{n}")).exists():
        n += 1
    return pasta / _nome_no_limite(pasta, nome, f"_{n}")


class SalvamentoInterrompido(RuntimeError):
    pass


def publicar_arquivo(tmp: Path, pasta: Path, nome: str, parar=None, avisar=None) -> Path:
    """Copia com criação exclusiva; nenhum rename/replace é necessário.

    Mantém o handle aberto durante a gravação. O nome final pode aparecer
    na pasta enquanto copia, mas só recebe status Baixado depois da conclusão.
    A fonte local permanece intacta em toda falha ou interrupção.
    """
    tmp, pasta = Path(tmp), Path(pasta)
    tamanho = tmp.stat().st_size
    if tamanho == 0:
        raise OSError("O arquivo local está vazio; nada foi salvo no destino.")

    def verificar_parada():
        if parar is not None and parar.is_set():
            raise SalvamentoInterrompido("Salvamento interrompido pelo usuário.")

    def aguardar(segundos):
        if parar is None:
            time.sleep(segundos)
        elif parar.wait(segundos):
            verificar_parada()

    tentativa = 0
    while True:
        verificar_parada()
        final = caminho_unico(pasta, nome)
        criado = False
        try:
            with tmp.open("rb") as origem, final.open("xb") as saida:
                criado = True
                while True:
                    verificar_parada()
                    bloco = origem.read(1024 * 1024)
                    if not bloco:
                        break
                    saida.write(bloco)
                saida.flush()
                os.fsync(saida.fileno())
                if saida.tell() != tamanho:
                    raise OSError("A cópia na pasta de destino ficou incompleta.")
            if final.stat().st_size != tamanho:
                raise OSError("O tamanho do arquivo salvo não corresponde ao PDF local.")
            return final
        except FileExistsError:
            # Outro processo ocupou o nome após a consulta; escolhe o próximo.
            continue
        except Exception as erro:
            if criado:
                try:
                    final.unlink(missing_ok=True)
                except OSError as limpeza:
                    raise OSError(f"Não foi possível remover a cópia incompleta em '{final}': {limpeza}. "
                                  "O arquivo não foi registrado como baixado.") from erro
            bloqueio = isinstance(erro, PermissionError) or getattr(erro, "winerror", None) in (5, 32, 33)
            if not bloqueio or tentativa >= len(ESPERAS_SALVAMENTO):
                raise
            espera = ESPERAS_SALVAMENTO[tentativa]
            tentativa += 1
            if avisar:
                avisar(f"Destino bloqueado ao salvar '{final.name}': {erro}. "
                       f"Nova tentativa {tentativa + 1}/{len(ESPERAS_SALVAMENTO) + 1} em {espera}s; "
                       "o PDF já baixado será reutilizado.")
            aguardar(espera)


def parse_linha(texto: str, guid: str = "") -> dict:
    """Interpreta o texto de uma linha da lista de acordos do DocuSign."""
    partes = [t.strip() for t in texto.split("\n") if t.strip()]
    uteis = [p for p in partes
             if p.lower() not in RUIDO and not re.fullmatch(r"\+\d+ mais", p)
             and not RE_DATA.match(p) and not RE_HORA.match(p)]
    para = next((p for p in uteis if p.startswith(("Para:", "De:"))), "")
    nome = next((p for p in uteis if not p.startswith(("Para:", "De:"))), "")
    data = next((p for p in partes if RE_DATA.match(p)), "")
    hora = next((p for p in partes if RE_HORA.match(p)), "")
    # Mantém a regra da versão funcional: primeiro destinatário que não seja Anderson.
    colab = ""
    if para:
        nomes = [re.sub(r"\s*\+\d+\s*mais\s*$", "", n, flags=re.I).strip()
                 for n in para.split(":", 1)[-1].split(",") if n.strip()]
        outros = [n for n in nomes if "anderson" not in n.lower()]
        colab = outros[0] if outros else ""
    return {
        "guid": guid,
        "nome": nome,
        "para": para,
        "colaborador": colab,
        "data": data,
        "hora": hora,
        "chave": guid or f"{nome}|{data}|{hora}",
    }


def caminho_log(destino: str) -> Path:
    return Path(destino) / "_log_baixados.csv"


def ler_log(destino: str) -> list:
    arq = caminho_log(destino)
    if not arq.exists():
        return []
    try:
        with open(arq, encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f, delimiter=";"))
    except Exception:
        return []


def gravar_log(destino: str, chave: str, arquivo: str):
    arq = caminho_log(destino)
    novo = not arq.exists()
    with open(arq, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        if novo:
            w.writerow(["chave", "arquivo", "baixado_em"])
        w.writerow([chave, arquivo, datetime.now().strftime("%d/%m/%Y %H:%M:%S")])


# ════════════════════════════════════════════════════════════════════
# WORKER PLAYWRIGHT (thread dedicada — o Playwright sync exige
# que todos os comandos rodem na mesma thread que o iniciou)
# ════════════════════════════════════════════════════════════════════
class PlaywrightWorker(threading.Thread):
    def __init__(self, ui_q: queue.Queue):
        super().__init__(daemon=True)
        self.ui = ui_q
        self.cmds = queue.Queue()
        self.parar = threading.Event()
        self.pw = None
        self.ctx = None
        self.page = None
        self.pasta_download = None

    def enviar(self, cmd, **kw):
        self.cmds.put((cmd, kw))

    def post(self, *msg):
        self.ui.put(msg)

    def log(self, txt):
        self.post("log", txt)

    def run(self):
        while True:
            cmd, kw = self.cmds.get()
            if cmd == "sair":
                self._fechar()
                break
            self.post("estado", "ocupado")
            try:
                getattr(self, f"_cmd_{cmd}")(**kw)
            except Exception as e:
                self.log(f"ERRO em '{cmd}': {e}")
            self.post("estado", "aberto" if self.page else "fechado")

    # ── página ativa (usa a aba que estiver em uso) ──
    def _pagina(self):
        if not self.ctx:
            return None
        abas = [p for p in self.ctx.pages if "docusign" in (p.url or "")]
        self.page = abas[-1] if abas else (self.ctx.pages[-1] if self.ctx.pages else None)
        return self.page

    # ── comandos ──
    def _cmd_abrir(self, url):
        if self.ctx:
            self._pagina().bring_to_front()
            self.log("Edge já está aberto.")
            return
        PERFIL_EDGE.mkdir(parents=True, exist_ok=True)
        self.pasta_download = PASTA_TEMP / f"_sessao_{uuid.uuid4().hex}"
        self.pasta_download.mkdir(parents=True, exist_ok=True)
        self.pw = sync_playwright().start()
        self.ctx = self.pw.chromium.launch_persistent_context(
            str(PERFIL_EDGE),
            channel="msedge",
            headless=False,
            accept_downloads=True,
            downloads_path=str(self.pasta_download),
            no_viewport=True,
            args=["--start-maximized"],
            ignore_default_args=["--no-sandbox"],
        )
        self.page = self.ctx.pages[0] if self.ctx.pages else self.ctx.new_page()
        self.page.goto(url)
        self.log("Versão 3.5.2. Edge aberto. Faça login (só na 1ª vez), ajuste filtros e clique em 'Ler envelopes'.")

    def _botoes_lista(self):
        return self.page.locator(SEL_BAIXAR_LISTA)

    def _rolar_lista(self):
        """Rola a lista para carregar itens sob demanda (lazy load)."""
        anterior, estavel = -1, 0
        self.page.mouse.move(800, 500)
        for _ in range(20):
            qtd = self._botoes_lista().count()
            if qtd == anterior:
                estavel += 1
                if estavel >= 2:
                    break
            else:
                estavel = 0
            anterior = qtd
            self.page.mouse.wheel(0, 4000)
            self.page.wait_for_timeout(700)
        self.page.mouse.wheel(0, -40000)

    def _cmd_ler(self, rolar):
        if not self._pagina():
            self.log("Abra o Edge primeiro.")
            return
        try:
            self.page.wait_for_selector(SEL_BAIXAR_LISTA, timeout=15000)
        except PWTimeout:
            self.log(f"Nenhum botão Baixar encontrado. URL atual: {self.page.url}")
            self.log("Use 'Diagnóstico' e me envie o resultado do Log.")
            self.post("linhas", [])
            return
        if rolar:
            self.log("Rolando a lista para carregar todos os itens...")
            self._rolar_lista()
        brutos = self.page.eval_on_selector_all(SEL_BAIXAR_LISTA, JS_LER_LINHAS)
        dados = [parse_linha(b["texto"], b["guid"]) for b in brutos]
        sem_texto = sum(1 for b in brutos if not b["texto"])
        if sem_texto:
            self.log(f"Aviso: {sem_texto} linha(s) sem texto legível (nome/data não capturados).")
        self.post("linhas", dados)
        self.log(f"{len(dados)} envelopes encontrados na página.")

    def _cmd_diagnostico(self):
        if not self._pagina():
            self.log("Abra o Edge primeiro.")
            return
        p = self.page
        self.log(f"URL: {p.url}")
        self.log(f"Abas abertas: {len(self.ctx.pages)}")
        self.log(f"Modal (abra antes de clicar): checkbox={p.locator(SEL_COMB_INPUT).count()}  "
                 f"rótulo={p.locator(SEL_COMB_LABEL).count()}  botão={p.locator(SEL_MODAL_BTN).count()}")
        self.log(f"Botões Baixar (data-qa): {p.locator(SEL_BAIXAR_LISTA).count()}")
        self.log(f"Elementos com texto 'Baixar': {p.get_by_text('Baixar', exact=True).count()}")
        self.log(f"role=row: {p.get_by_role('row').count()}  |  role=dialog: {p.get_by_role('dialog').count()}")
        qas = p.eval_on_selector_all(
            "[data-qa]",
            "els => els.filter(e => /baixar|download/i.test(e.innerText||'') || /download/i.test(e.getAttribute('data-qa')))"
            ".slice(0,15).map(e => e.tagName + ' | ' + e.getAttribute('data-qa'))")
        for q in qas:
            self.log(f"  data-qa: {q}")

    def _botao_modal(self):
        """Localiza o botão 'Baixar' do modal 'Fazer download dos arquivos'."""
        p = self.page
        candidatos = [
            p.locator('[role="dialog"], [aria-modal="true"]').filter(has_text="Fazer download"),
            p.locator('[role="dialog"], [aria-modal="true"]'),
            p.locator("div").filter(has_text="Fazer download dos arquivos").filter(
                has=p.get_by_role("button", name="Baixar", exact=True)),
        ]
        for modal in candidatos:
            try:
                modal.last.wait_for(state="visible", timeout=4000)
            except PWTimeout:
                continue
            m = modal.last
            btn = m.get_by_role("button", name="Baixar", exact=True)
            if btn.count():
                return m, btn.last
        raise PWTimeout("Modal de download não encontrado")

    def _garantir_combinar(self, combinar):
        """Lê o checkbox real 'Combinar todos os PDFs', ajusta só se precisar e confirma.
        Retorna (confirmado: bool, mensagem)."""
        p = self.page
        inp = p.locator(SEL_COMB_INPUT)
        lbl = p.locator(SEL_COMB_LABEL)
        try:
            inp.first.wait_for(state="attached", timeout=5000)
        except PWTimeout:
            return (False,
                    f"checkbox não encontrado (input={inp.count()}, rótulo={lbl.count()})")

        def ler():
            try:
                return bool(inp.first.evaluate("e => e.checked"))
            except Exception:
                return None

        antes = ler()
        if antes is None:
            return False, "não consegui ler o estado do checkbox"
        if antes == combinar:
            return True, f"já estava {'marcado' if antes else 'desmarcado'} (não cliquei)"

        for _ in range(2):
            try:
                if lbl.count():
                    lbl.first.click(timeout=3000)
                else:
                    inp.first.click(force=True, timeout=3000)
            except Exception as e:
                self.log(f"  clique no checkbox falhou: {str(e).splitlines()[0]}")
            p.wait_for_timeout(300)
            agora = ler()
            if agora == combinar:
                return True, (f"estava {'marcado' if antes else 'desmarcado'} -> "
                              f"{'marcado' if agora else 'desmarcado'} (confirmado)")
        return False, f"não confirmou após clicar (estado atual: {ler()})"

    def _clicar_e_aguardar_download(self, btn):
        """Espera a transferência completa sem bloquear em Download.save_as.

        O Edge (Chromium) grava GUID.crdownload durante a transferência e
        renomeia para GUID ao concluir. A pasta é exclusiva desta sessão e
        configurada pela opção pública downloads_path do Playwright.
        Todos os comandos continuam na mesma thread do Playwright sync.
        """
        if self.parar.is_set():
            return None, "interrompido"
        if not self.pasta_download:
            raise RuntimeError("Pasta temporária da sessão indisponível. Feche e reabra o Edge pelo programa.")
        page = self.page
        recebido = []
        anteriores = {p.name for p in self.pasta_download.iterdir()}
        concluido = False

        def on_download(d):
            recebido.append(d)

        page.on("download", on_download)
        try:
            inicio = time.monotonic()
            limite = inicio + ESPERA_DOWNLOAD_S
            inicio_transferencia = None
            btn.click(timeout=3000, no_wait_after=True)
            aviso = 0
            while time.monotonic() < limite:
                if self.parar.is_set():
                    return None, "interrompido"
                if len(recebido) > 1:
                    raise RuntimeError("Mais de um download iniciado; não foi possível identificar o envelope.")
                if recebido and inicio_transferencia is None:
                    inicio_transferencia = time.monotonic()
                    limite = inicio_transferencia + ESPERA_TRANSFERENCIA_S
                novos = [p for p in self.pasta_download.iterdir() if p.is_file() and p.name not in anteriores]
                pendentes = [p for p in novos if p.name.endswith(".crdownload")]
                completos = [p for p in novos if re.fullmatch(
                    r"[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}", p.name)]
                if recebido and not pendentes and completos:
                    if len(completos) != 1:
                        raise RuntimeError("Mais de um arquivo na transferência; envelope não identificado.")
                    concluido = True
                    return recebido[0], ""
                decorrido = int(time.monotonic() - inicio)
                if decorrido // 15 > aviso:
                    aviso = decorrido // 15
                    self.log(f"  aguardando o arquivo completo... {decorrido}s")
                page.wait_for_timeout(100)
            if inicio_transferencia is None:
                return None, f"download não iniciou em {ESPERA_DOWNLOAD_S}s"
            return None, f"transferência não foi concluída em {ESPERA_TRANSFERENCIA_S}s"
        finally:
            try:
                page.remove_listener("download", on_download)
            except Exception:
                pass
            if not concluido:
                for d in recebido:
                    try:
                        d.cancel()
                    except Exception:
                        pass

    def _cmd_baixar(self, itens, destino, intervalo, combinar):
        if not self._pagina():
            self.log("Abra o Edge primeiro.")
            return
        PASTA_TEMP.mkdir(parents=True, exist_ok=True)
        pasta_dest = Path(destino)
        pasta_dest.mkdir(parents=True, exist_ok=True)
        page = self.page
        ok = erro = 0

        for n, it in enumerate(itens, 1):
            if self.parar.is_set():
                self.log("Interrompido pelo usuário.")
                break
            chave = it["chave"]
            self.post("status", chave, "Baixando...", "")
            self.post("progresso", n - 1, len(itens))

            alvo = page.locator(f'[data-qa="{PREFIXO_ROW}{it["guid"]}{SUFIXO_DL}"]') if it["guid"] else None
            if alvo is None or alvo.count() == 0:
                self.post("status", chave, "Não encontrado", "")
                self.log(f"Não encontrado na tela: {it['nome']}")
                erro += 1
                continue
            try:
                alvo.first.scroll_into_view_if_needed()
                alvo.first.click()
                modal, btn = self._botao_modal()

                ok_comb, msg_comb = self._garantir_combinar(combinar)
                self.log(f"Combinar PDFs: {msg_comb}")
                if not ok_comb:
                    raise RuntimeError("opção 'Combinar PDFs' não confirmada — envelope não baixado")

                btn_real = page.locator(SEL_MODAL_BTN)
                if btn_real.count():
                    btn = btn_real.last

                dl, motivo = self._clicar_e_aguardar_download(btn)
                if dl is None:
                    if motivo == "interrompido":
                        self.post("status", chave, "Pendente", "")
                        self.log(f"Parado: '{it['nome']}' não foi baixado.")
                        try:
                            page.keyboard.press("Escape")
                        except Exception:
                            pass
                        break
                    raise RuntimeError(motivo)

                self.log(f"Arquivo recebido do DocuSign: {dl.suggested_filename}")
                ext = Path(dl.suggested_filename).suffix or ".pdf"
                if combinar and ext.lower() != ".pdf":
                    self.log(f"Aviso: com 'Combinar' confirmado, o DocuSign entregou '{ext}'. "
                             f"Envie esta linha do Log para análise.")
                nome_arq = montar_nome_arquivo(it, dl.suggested_filename)
                # Nome curto local: o título só é aplicado na pasta de destino.
                tmp = PASTA_TEMP / f"{uuid.uuid4().hex}{ext}"
                if self.parar.is_set():
                    try:
                        dl.cancel()
                    except Exception:
                        pass
                    self.post("status", chave, "Pendente", "")
                    self.log(f"Parado: download de '{it['nome']}' cancelado.")
                    break
                try:
                    dl.save_as(str(tmp))
                    falha = dl.failure()
                    if falha:
                        raise RuntimeError(f"download não concluído: {falha}")
                    if not tmp.is_file() or tmp.stat().st_size == 0:
                        raise RuntimeError("O DocuSign entregou um arquivo vazio.")
                except Exception:
                    try:
                        tmp.unlink(missing_ok=True)
                    except OSError:
                        pass
                    raise
                try:
                    final = publicar_arquivo(tmp, pasta_dest, nome_arq, parar=self.parar, avisar=self.log)
                except SalvamentoInterrompido:
                    self.post("status", chave, "Pendente", "")
                    self.log(f"Parado durante o salvamento de '{it['nome']}'. Cópia completa preservada em: {tmp}")
                    try:
                        page.keyboard.press("Escape")
                    except Exception:
                        pass
                    break
                except Exception as e:
                    raise RuntimeError(f"Falha ao salvar '{nome_arq}' em '{pasta_dest}': {e}. "
                                       f"Cópia completa preservada em: {tmp}") from e
                try:
                    tmp.unlink()
                except OSError as e:
                    self.log(f"Arquivo salvo; cópia temporária mantida em '{tmp}': {e}")
                try:
                    gravar_log(destino, chave, final.name)
                except OSError as e:
                    self.log(f"AVISO: arquivo salvo em '{final}', mas o histórico não pôde ser gravado: {e}. "
                             "Confira este arquivo antes de baixar o envelope novamente.")
                self.post("status", chave, "Baixado", final.name)
                self.log(f"OK: {final}")
                ok += 1

                try:
                    modal.wait_for(state="hidden", timeout=5000)
                except Exception:
                    try:
                        page.keyboard.press("Escape")
                    except Exception:
                        pass
            except Exception as e:
                self.post("status", chave, "Erro", "")
                self.log(f"ERRO em '{it['nome']}': {str(e).splitlines()[0]}")
                erro += 1
                try:
                    page.keyboard.press("Escape")
                except Exception:
                    pass
            for _ in range(max(1, int(intervalo * 10))):
                if self.parar.is_set():
                    break
                page.wait_for_timeout(100)

        if self.parar.is_set():
            self.log("Interrompido pelo usuário. Arquivos já concluídos foram mantidos.")
        self.post("progresso", ok + erro, max(len(itens), 1))
        self.post("fim", ok, erro, self.parar.is_set())
        self.log(f"{'Interrompido' if self.parar.is_set() else 'Concluído'}: {ok} baixados, {erro} com erro.")

    def _fechar(self):
        try:
            if self.ctx:
                self.ctx.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass
        self.ctx = self.page = self.pw = None
        if self.pasta_download:
            shutil.rmtree(self.pasta_download, ignore_errors=True)
            self.pasta_download = None

    def _cmd_fechar(self):
        self._fechar()
        self.log("Edge da automação fechado.")


# ════════════════════════════════════════════════════════════════════
# COMPONENTES DE UI
# ════════════════════════════════════════════════════════════════════
def configurar_estilos(root):
    st = ttk.Style(root)
    st.theme_use("clam")
    for orient, nome in (("Vertical", "Slim.Vertical.TScrollbar"), ("Horizontal", "Slim.Horizontal.TScrollbar")):
        st.layout(nome, [(f"{orient}.Scrollbar.trough", {
            "sticky": "ns" if orient == "Vertical" else "ew",
            "children": [(f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})],
        })])
        st.configure(nome, troughcolor=TRACK, background=THUMB, bordercolor=TRACK,
                     lightcolor=THUMB, darkcolor=THUMB, gripcount=0, width=11, arrowsize=11)
        st.map(nome, background=[("active", AZUL), ("pressed", AZUL)],
               lightcolor=[("active", AZUL)], darkcolor=[("active", AZUL)])
    st.configure("OFC.Treeview", background=BRANCO, fieldbackground=BRANCO, foreground=TXT,
                 rowheight=24, borderwidth=0, font=(FONTE_CORPO, 9))
    st.configure("OFC.Treeview.Heading", background=TH_BG, foreground=AZUL_ESC,
                 font=(FONTE_CORPO, 9, "bold"), relief="flat", borderwidth=0, padding=(6, 4))
    st.map("OFC.Treeview.Heading", background=[("active", "#D3E7F4")])
    st.map("OFC.Treeview", background=[("selected", AZUL_CLARO)], foreground=[("selected", BRANCO)])


def scrollbar_autohide(sb):
    def _set(lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            sb.grid_remove()
        else:
            sb.grid()
        sb.set(lo, hi)
    return _set


class AbaRolavel(tk.Frame):
    """Canvas rolável por aba, com scrollbar fina auto-ocultável."""

    def __init__(self, master):
        super().__init__(master, bg=BRANCO)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(self, bg=BRANCO, highlightthickness=0)
        self.sb = ttk.Scrollbar(self, orient="vertical", style="Slim.Vertical.TScrollbar",
                                command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar_autohide(self.sb))
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.sb.grid(row=0, column=1, sticky="ns")
        self.interno = tk.Frame(self.canvas, bg=BRANCO)
        self._win = self.canvas.create_window((0, 0), window=self.interno, anchor="nw")
        self.interno.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._roda))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _roda(self, e):
        w = e.widget
        if isinstance(w, (ttk.Treeview, tk.Text)) or "ctktextbox" in str(w).lower():
            return
        if self.sb.winfo_ismapped():
            self.canvas.yview_scroll(int(-e.delta / 120), "units")


class TabelaOFC(tk.Frame):
    """Treeview com borda #CCCCCC, ordenação por clique (^/v) e zebra recalculada."""

    def __init__(self, master, colunas, altura=10, on_duplo=None):
        super().__init__(master, bg=BRANCO, highlightbackground=BORDA, highlightthickness=1)
        self.colunas = colunas
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        ids = [c[0] for c in colunas]
        self.tree = ttk.Treeview(self, columns=ids, show="headings", height=altura,
                                 style="OFC.Treeview", selectmode="extended")
        vsb = ttk.Scrollbar(self, orient="vertical", style="Slim.Vertical.TScrollbar", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", style="Slim.Horizontal.TScrollbar", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrollbar_autohide(vsb), xscrollcommand=scrollbar_autohide(hsb))
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.ordem = {"col": None, "desc": False}
        for cid, tit, larg, elastica, numerica in colunas:
            self.tree.heading(cid, text=tit, anchor="w", command=lambda c=cid: self.ordenar(c))
            self.tree.column(cid, width=larg, minwidth=60, stretch=elastica, anchor="e" if numerica else "w")
        self.tree.tag_configure("par", background=BRANCO)
        self.tree.tag_configure("impar", background=ZEBRA)
        self.tree.tag_configure("Baixado", foreground=VERDE)
        self.tree.tag_configure("Erro", foreground=VINHO)
        self.tree.tag_configure("Não encontrado", foreground=VINHO)
        self.tree.tag_configure("Já baixado", foreground=CINZA)
        self.tree.tag_configure("Fora do filtro", foreground=CINZA)
        self.tree.tag_configure("Baixando...", foreground=AZUL)
        if on_duplo:
            self.tree.bind("<Double-1>", lambda e: self._duplo(on_duplo))

    def _duplo(self, cb):
        sel = self.tree.focus()
        if sel:
            cb(sel, dict(zip([c[1] for c in self.colunas], self.tree.item(sel, "values"))))

    def limpar(self):
        self.tree.delete(*self.tree.get_children())

    def inserir(self, iid, valores, status_tag=""):
        self.tree.insert("", "end", iid=iid, values=valores, tags=(status_tag,))

    def atualizar(self, iid, coluna, valor, status_tag=None):
        if not self.tree.exists(iid):
            return
        self.tree.set(iid, coluna, valor)
        if status_tag is not None:
            zeb = [t for t in self.tree.item(iid, "tags") if t in ("par", "impar")]
            self.tree.item(iid, tags=tuple(zeb) + (status_tag,))

    def zebra(self):
        for i, iid in enumerate(self.tree.get_children()):
            status = [t for t in self.tree.item(iid, "tags") if t not in ("par", "impar")]
            self.tree.item(iid, tags=("par" if i % 2 == 0 else "impar", *status))

    def ordenar(self, col, forcar_asc=False):
        desc = False if forcar_asc or self.ordem["col"] != col else not self.ordem["desc"]
        self.ordem = {"col": col, "desc": desc}

        def chave(v):
            if RE_DATA.match(v or ""):
                d, m, a = v.split("/")
                return (0, f"{a}{m}{d}")
            m = re.match(r"^(\d{2})/(\d{2})/(\d{4}) (\d{2}:\d{2})", v or "")
            if m:
                return (0, f"{m[3]}{m[2]}{m[1]}{m[4]}")
            try:
                return (1, float(str(v).replace(".", "").replace(",", ".")))
            except ValueError:
                return (2, str(v).lower())

        itens = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        itens.sort(key=lambda t: chave(t[0]), reverse=desc)
        for idx, (_, k) in enumerate(itens):
            self.tree.move(k, "", idx)
        for cid, tit, *_ in self.colunas:
            seta = (" v" if desc else " ^") if cid == col else ""
            self.tree.heading(cid, text=tit + seta)
        self.zebra()

    def selecionados(self):
        return list(self.tree.selection())

    def todos(self):
        return list(self.tree.get_children())


def botao(master, texto, tipo, comando, largura=150):
    cores = {
        "avancar": (VERDE, VERDE_HOVER, BRANCO),
        "primario": (AZUL, AZUL_ESC, BRANCO),
        "secundario": (AZUL_CLARO, AZUL, BRANCO),
        "auxiliar": (CINZA, CINZA_HOVER, BRANCO),
        "destrutivo": (VERMELHO, VERMELHO_HOVER, BRANCO),
    }[tipo]
    return ctk.CTkButton(master, text=texto, command=comando, width=largura, height=32,
                         fg_color=cores[0], hover_color=cores[1], text_color=cores[2],
                         corner_radius=4, font=ctk.CTkFont(FONTE_CORPO, 12, "bold"))


# ════════════════════════════════════════════════════════════════════
# APLICAÇÃO
# ════════════════════════════════════════════════════════════════════
class App(ctk.CTk):
    COLS_ENV = [
        ("status", "Status", 110, False, False),
        ("nome", "Envelope", 340, True, False),
        ("colaborador", "Colaborador", 230, True, False),
        ("para", "Para / De", 260, True, False),
        ("data", "Data", 95, False, True),
        ("hora", "Hora", 65, False, True),
        ("arquivo", "Arquivo salvo", 280, True, False),
        ("guid", "Envelope ID", 260, False, False),
    ]
    COLS_HIST = [
        ("baixado_em", "Baixado em", 150, False, True),
        ("arquivo", "Arquivo", 420, True, False),
        ("chave", "Envelope ID / chave", 380, True, False),
    ]

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title(APP_NOME)
        self.geometry("1280x820")
        self.minsize(1050, 650)
        self.configure(fg_color=BRANCO)
        configurar_estilos(self)

        self.cfg = carregar_config()
        self.ui_q = queue.Queue()
        self.worker = None
        self.envelopes = {}
        self.iid_por_chave = {}
        self.baixados_sessao = 0
        self.erros_sessao = 0

        self._montar_header()
        self._montar_footer()
        self._montar_abas()
        self._carregar_historico()

        if PLAYWRIGHT_OK:
            self.worker = PlaywrightWorker(self.ui_q)
            self.worker.start()
        else:
            self._log("Playwright não encontrado. Rode no PowerShell:  python.exe -m pip install playwright")
            messagebox.showwarning(APP_NOME, "Playwright não instalado.\n\nRode no PowerShell:\n"
                                             "python.exe -m pip install playwright")
        self._estado("fechado")
        self.after(100, self._processar_fila)
        self.protocol("WM_DELETE_WINDOW", self._sair)

    def _montar_header(self):
        h = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=56)
        h.pack(fill="x")
        h.pack_propagate(False)
        ctk.CTkLabel(h, text=APP_NOME, text_color=BRANCO,
                     font=ctk.CTkFont(FONTE_TIT, 20, "bold")).pack(side="left", padx=18)
        self.lbl_estado = ctk.CTkLabel(h, text="", text_color=BRANCO, font=ctk.CTkFont(FONTE_CORPO, 12))
        self.lbl_estado.pack(side="right", padx=18)

    def _montar_footer(self):
        f = ctk.CTkFrame(self, fg_color=AZUL, corner_radius=0, height=26)
        f.pack(side="bottom", fill="x")
        f.pack_propagate(False)
        ctk.CTkLabel(f, text="Anderson Souza  |  RH Sonova Brasil", text_color=BRANCO,
                     font=ctk.CTkFont(FONTE_CORPO, 11)).pack(side="right", padx=14)
        self.lbl_rodape = ctk.CTkLabel(f, text="", text_color=BRANCO, font=ctk.CTkFont(FONTE_CORPO, 11))
        self.lbl_rodape.pack(side="left", padx=14)

    def _montar_abas(self):
        self.tabs = ctk.CTkTabview(self, fg_color=BRANCO, segmented_button_selected_color=AZUL,
                                   segmented_button_selected_hover_color=AZUL_ESC,
                                   segmented_button_unselected_color=TH_BG,
                                   segmented_button_unselected_hover_color=AZUL_CLARO,
                                   text_color=AZUL_ESC, border_color=BORDA, border_width=1)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(8, 8))
        for nome in ("Envelopes", "Histórico", "Log"):
            self.tabs.add(nome)
        self._aba_envelopes(self._rolavel("Envelopes"))
        self._aba_historico(self._rolavel("Histórico"))
        self._aba_log(self.tabs.tab("Log"))

    def _rolavel(self, aba):
        r = AbaRolavel(self.tabs.tab(aba))
        r.pack(fill="both", expand=True)
        return r.interno

    def _card(self, master, titulo):
        c = ctk.CTkFrame(master, fg_color=BRANCO, border_color=BORDA, border_width=1, corner_radius=6)
        c.pack(fill="x", padx=6, pady=6)
        ctk.CTkLabel(c, text=titulo, text_color=AZUL_ESC,
                     font=ctk.CTkFont(FONTE_TIT, 14, "bold")).pack(anchor="w", padx=12, pady=(8, 2))
        corpo = ctk.CTkFrame(c, fg_color=BRANCO)
        corpo.pack(fill="x", padx=12, pady=(0, 10))
        return corpo

    def _aba_envelopes(self, m):
        cfg = self._card(m, "Configuração")
        cfg.columnconfigure(1, weight=1)
        fonte = ctk.CTkFont(FONTE_CORPO, 12)

        def campo(linha, rotulo, valor, com_botao=None):
            ctk.CTkLabel(cfg, text=rotulo, text_color=TXT, font=fonte).grid(row=linha, column=0, sticky="w", pady=3)
            e = ctk.CTkEntry(cfg, font=fonte, border_color=BORDA, text_color=TXT)
            e.insert(0, str(valor))
            e.grid(row=linha, column=1, sticky="ew", padx=(10, 6), pady=3)
            if com_botao:
                com_botao.grid(row=linha, column=2, pady=3)
            return e

        self.ent_url = campo(0, "URL DocuSign", self.cfg["url"])
        self.ent_dest = campo(1, "Pasta destino", self.cfg["destino"],
                              botao(cfg, "Escolher...", "auxiliar", self._escolher_pasta, 110))
        self.ent_filtro = campo(2, "Filtro (regex no nome)", self.cfg["filtro"])
        self.ent_interv = campo(3, "Intervalo entre downloads (s)", self.cfg["intervalo"])

        opc = ctk.CTkFrame(cfg, fg_color=BRANCO)
        opc.grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.var_comb = tk.BooleanVar(value=self.cfg["combinar"])
        self.var_ign = tk.BooleanVar(value=self.cfg["ignorar_baixados"])
        self.var_rolar = tk.BooleanVar(value=True)
        self.chk_comb = ctk.CTkCheckBox(
            opc, text="Combinar PDFs em um único arquivo", variable=self.var_comb,
            command=self._alterar_modo_download, font=fonte, text_color=TXT,
            fg_color=AZUL, hover_color=AZUL_ESC, border_color=BORDA)
        self.chk_comb.pack(side="left", padx=(0, 22))
        for txt, var in (("Ignorar envelopes já baixados (log)", self.var_ign),
                         ("Rolar a lista antes de ler (carregar tudo)", self.var_rolar)):
            ctk.CTkCheckBox(opc, text=txt, variable=var, font=fonte, text_color=TXT, fg_color=AZUL,
                            hover_color=AZUL_ESC, border_color=BORDA).pack(side="left", padx=(0, 22))
        self.lbl_modo_download = ctk.CTkLabel(
            cfg, text=texto_modo_download(self.var_comb.get()), font=fonte,
            text_color=TXT2, anchor="w")
        self.lbl_modo_download.grid(row=5, column=0, columnspan=3, sticky="w", pady=(2, 0))

        # Barra de ações em duas linhas (fluxo principal / auxiliares)
        barra1 = ctk.CTkFrame(m, fg_color=BRANCO)
        barra1.pack(fill="x", padx=6, pady=(4, 0))
        barra2 = ctk.CTkFrame(m, fg_color=BRANCO)
        barra2.pack(fill="x", padx=6, pady=(0, 2))
        self.btn_abrir = botao(barra1, "1. Abrir Edge", "primario", self._abrir)
        self.btn_ler = botao(barra1, "2. Ler envelopes", "primario", self._ler)
        self.btn_pend = botao(barra1, "3. Baixar pendentes", "avancar", lambda: self._baixar(False), 170)
        self.btn_sel = botao(barra1, "Baixar selecionados", "secundario", lambda: self._baixar(True), 170)
        self.btn_parar = botao(barra1, "Parar", "destrutivo", self._parar, 90)
        for b in (self.btn_abrir, self.btn_ler, self.btn_pend, self.btn_sel, self.btn_parar):
            b.pack(side="left", padx=(0, 6), pady=4)
        self.btn_pasta = botao(barra2, "Abrir pasta", "auxiliar", self._abrir_pasta, 110)
        self.btn_diag = botao(barra2, "Diagnóstico", "auxiliar", self._diagnostico, 120)
        self.btn_codegen = botao(barra2, "Codegen (seletores)", "auxiliar", self._codegen, 160)
        self.btn_csv = botao(barra2, "Exportar CSV", "auxiliar", self._exportar_csv, 120)
        self.btn_fechar = botao(barra2, "Fechar Edge", "destrutivo", self._fechar_edge, 110)
        for b in (self.btn_pasta, self.btn_diag, self.btn_codegen, self.btn_csv, self.btn_fechar):
            b.pack(side="left", padx=(0, 6), pady=4)

        kp = ctk.CTkFrame(m, fg_color=BRANCO)
        kp.pack(fill="x", padx=6, pady=4)
        self.kpi = {}
        for i, (k, tit, cor) in enumerate((("pagina", "Na página", AZUL_ESC), ("filtrados", "No filtro", AZUL),
                                            ("pendentes", "Pendentes", VINHO), ("baixados", "Baixados agora", VERDE),
                                            ("erros", "Erros", TERRA))):
            kp.columnconfigure(i, weight=1)
            c = ctk.CTkFrame(kp, fg_color=BRANCO, border_color=BORDA, border_width=1, corner_radius=6)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 6, 0))
            ctk.CTkLabel(c, text=tit, text_color=TXT2, font=ctk.CTkFont(FONTE_CORPO, 11)).pack(pady=(8, 0))
            lbl = ctk.CTkLabel(c, text="0", text_color=cor, font=ctk.CTkFont(FONTE_TIT, 24, "bold"))
            lbl.pack(pady=(0, 8))
            self.kpi[k] = lbl

        self.prog = ctk.CTkProgressBar(m, progress_color=AZUL, fg_color=TRACK, height=8)
        self.prog.set(0)
        self.prog.pack(fill="x", padx=6, pady=(6, 4))

        self.tab_env = TabelaOFC(m, self.COLS_ENV, altura=12, on_duplo=self._detalhe)
        self.tab_env.pack(fill="both", expand=True, padx=6, pady=(4, 10))

    def _aba_historico(self, m):
        barra = ctk.CTkFrame(m, fg_color=BRANCO)
        barra.pack(fill="x", padx=6, pady=(6, 2))
        botao(barra, "Recarregar histórico", "primario", self._carregar_historico, 170).pack(side="left", padx=(0, 6))
        self.lbl_hist = ctk.CTkLabel(barra, text="", text_color=TXT2, font=ctk.CTkFont(FONTE_CORPO, 12))
        self.lbl_hist.pack(side="left", padx=10)
        self.tab_hist = TabelaOFC(m, self.COLS_HIST, altura=16, on_duplo=self._detalhe)
        self.tab_hist.pack(fill="both", expand=True, padx=6, pady=(4, 10))

    def _aba_log(self, m):
        self.txt_log = ctk.CTkTextbox(m, font=ctk.CTkFont("Consolas", 11), text_color=TXT,
                                      fg_color=BRANCO, border_color=BORDA, border_width=1)
        self.txt_log.pack(fill="both", expand=True, padx=6, pady=6)

    def _log(self, txt):
        linha = f"[{datetime.now().strftime('%H:%M:%S')}] {txt}\n"
        self.txt_log.insert("end", linha)
        self.txt_log.see("end")
        self.lbl_rodape.configure(text=txt[:110])

    def _estado(self, est):
        rot = {"fechado": "Edge: fechado", "aberto": "Edge: aberto  |  pronto",
               "ocupado": "Edge: trabalhando..."}[est]
        self.lbl_estado.configure(text=rot)
        ocupado = est == "ocupado"
        aberto = est == "aberto"
        self.btn_abrir.configure(state="disabled" if ocupado or not PLAYWRIGHT_OK else "normal")
        for b in (self.btn_ler, self.btn_pend, self.btn_sel, self.btn_fechar, self.btn_diag):
            b.configure(state="normal" if aberto else "disabled")
        self.btn_parar.configure(state="normal" if ocupado else "disabled")

    def _alterar_modo_download(self):
        cfg = self._ler_cfg()
        self.lbl_modo_download.configure(text=texto_modo_download(cfg["combinar"]))

    def _ler_cfg(self):
        try:
            interv = max(1.0, float(self.ent_interv.get().replace(",", ".")))
        except ValueError:
            interv = 3.0
        url = url_com_50_envelopes(self.ent_url.get())
        if url != self.ent_url.get().strip():
            self.ent_url.delete(0, "end")
            self.ent_url.insert(0, url)
        self.cfg.update({
            "url": url,
            "destino": self.ent_dest.get().strip(),
            "filtro": self.ent_filtro.get().strip(),
            "intervalo": interv,
            "combinar": self.var_comb.get(),
            "ignorar_baixados": self.var_ign.get(),
        })
        salvar_config(self.cfg)
        return self.cfg

    def _atualizar_kpis(self):
        vals = [self.tab_env.tree.set(i, "status") for i in self.tab_env.todos()]
        self.kpi["pagina"].configure(text=str(len(vals)))
        self.kpi["filtrados"].configure(text=str(sum(v != "Fora do filtro" for v in vals)))
        self.kpi["pendentes"].configure(text=str(sum(v in ("Pendente", "Erro", "Não encontrado") for v in vals)))
        self.kpi["baixados"].configure(text=str(self.baixados_sessao))
        self.kpi["erros"].configure(text=str(self.erros_sessao))

    def _escolher_pasta(self):
        p = filedialog.askdirectory(initialdir=self.ent_dest.get() or str(Path.home()))
        if p:
            self.ent_dest.delete(0, "end")
            self.ent_dest.insert(0, str(Path(p)))

    def _abrir(self):
        cfg = self._ler_cfg()
        self.worker.enviar("abrir", url=cfg["url"])

    def _ler(self):
        self._ler_cfg()
        self.worker.enviar("ler", rolar=self.var_rolar.get())

    def _diagnostico(self):
        self.tabs.set("Log")
        self.worker.enviar("diagnostico")

    def _baixar(self, so_selecionados):
        cfg = self._ler_cfg()
        iids = self.tab_env.selecionados() if so_selecionados else self.tab_env.todos()
        if not iids:
            messagebox.showinfo(APP_NOME, "Nenhum envelope para baixar. Leia os envelopes primeiro.")
            return
        mapa = {v: k for k, v in self.iid_por_chave.items()}
        itens = []
        for iid in iids:
            st = self.tab_env.tree.set(iid, "status")
            if not so_selecionados and st not in ("Pendente", "Erro", "Não encontrado"):
                continue
            chave = mapa.get(iid)
            if chave:
                itens.append(self.envelopes[chave])
        if not itens:
            messagebox.showinfo(APP_NOME, "Nenhum envelope pendente.")
            return
        if not messagebox.askyesno(APP_NOME, f"Baixar {len(itens)} envelope(s) para:\n\n{cfg['destino']}?"):
            return
        self.prog.set(0)
        self.worker.parar.clear()
        self.worker.enviar("baixar", itens=itens, destino=cfg["destino"],
                           intervalo=cfg["intervalo"], combinar=cfg["combinar"])

    def _parar(self):
        if self.worker:
            self.worker.parar.set()
            self._log("Parada solicitada — interrompendo a espera pela transferência. Arquivos já salvos serão mantidos.")

    def _fechar_edge(self):
        self.worker.enviar("fechar")

    def _abrir_pasta(self):
        p = self.ent_dest.get().strip()
        if os.path.isdir(p):
            os.startfile(p)
        else:
            messagebox.showwarning(APP_NOME, "Pasta não encontrada.")

    def _codegen(self):
        try:
            subprocess.Popen([sys.executable, "-m", "playwright", "codegen", "--channel", "msedge",
                              self.ent_url.get().strip()])
            self._log("Codegen aberto: clique em 'Baixar' e no 'Baixar' do modal para ver os seletores reais.")
        except Exception as e:
            self._log(f"Não consegui abrir o codegen: {e}")

    def _exportar_csv(self):
        arq = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                           initialfile=f"docusign_envelopes_{datetime.now():%Y%m%d_%H%M}.csv")
        if not arq:
            return
        with open(arq, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow([c[1] for c in self.COLS_ENV])
            for iid in self.tab_env.todos():
                w.writerow(self.tab_env.tree.item(iid, "values"))
        self._log(f"CSV exportado: {arq}")

    def _detalhe(self, iid, dados):
        win = ctk.CTkToplevel(self)
        win.title("Detalhe")
        win.geometry("680x420")
        win.configure(fg_color=BRANCO)
        win.transient(self)
        win.grab_set()
        cab = ctk.CTkFrame(win, fg_color=AZUL, corner_radius=0, height=40)
        cab.pack(fill="x")
        ctk.CTkLabel(cab, text="Detalhe do registro", text_color=BRANCO,
                     font=ctk.CTkFont(FONTE_TIT, 15, "bold")).pack(side="left", padx=14, pady=6)
        corpo = ctk.CTkFrame(win, fg_color=BRANCO)
        corpo.pack(fill="both", expand=True, padx=14, pady=10)
        corpo.columnconfigure(1, weight=1)
        for i, (k, v) in enumerate(dados.items()):
            ctk.CTkLabel(corpo, text=k, text_color=TXT2,
                         font=ctk.CTkFont(FONTE_CORPO, 12, "bold")).grid(row=i, column=0, sticky="nw", pady=3)
            e = ctk.CTkTextbox(corpo, height=28, font=ctk.CTkFont(FONTE_CORPO, 12), text_color=TXT,
                               fg_color=BRANCO, border_color=BORDA, border_width=1, wrap="word")
            e.insert("1.0", str(v))
            e.configure(state="disabled")
            e.grid(row=i, column=1, sticky="ew", padx=(10, 0), pady=3)
        botao(win, "Fechar", "auxiliar", win.destroy, 100).pack(pady=(0, 12))

    def _carregar_historico(self):
        destino = self.ent_dest.get().strip() if hasattr(self, "ent_dest") else self.cfg["destino"]
        regs = ler_log(destino)
        self.tab_hist.limpar()
        for i, r in enumerate(regs):
            self.tab_hist.inserir(f"h{i}", (r.get("baixado_em", ""), r.get("arquivo", ""), r.get("chave", "")))
        self.tab_hist.ordenar("baixado_em", forcar_asc=True)
        self.lbl_hist.configure(text=f"{len(regs)} registro(s) em {caminho_log(destino).name}")
        return {r.get("chave") for r in regs}

    def _processar_fila(self):
        try:
            while True:
                msg = self.ui_q.get_nowait()
                tipo = msg[0]
                if tipo == "log":
                    self._log(msg[1])
                elif tipo == "estado":
                    self._estado(msg[1])
                elif tipo == "linhas":
                    self._receber_linhas(msg[1])
                elif tipo == "status":
                    _, chave, status, arquivo = msg
                    iid = self.iid_por_chave.get(chave)
                    if iid:
                        self.tab_env.atualizar(iid, "status", status, status)
                        if arquivo:
                            self.tab_env.atualizar(iid, "arquivo", arquivo)
                    if status == "Baixado":
                        self.baixados_sessao += 1
                    elif status in ("Erro", "Não encontrado"):
                        self.erros_sessao += 1
                    self._atualizar_kpis()
                elif tipo == "progresso":
                    _, feito, total = msg
                    self.prog.set(feito / total if total else 0)
                elif tipo == "fim":
                    self._carregar_historico()
                    titulo = "Interrompido" if msg[3] else "Concluído"
                    messagebox.showinfo(APP_NOME, f"{titulo}.\n\nBaixados: {msg[1]}\nErros: {msg[2]}")
        except queue.Empty:
            pass
        self.after(100, self._processar_fila)

    def _receber_linhas(self, linhas):
        cfg = self._ler_cfg()
        try:
            filtro = re.compile(cfg["filtro"], re.I) if cfg["filtro"] else None
        except re.error:
            filtro = None
            self._log("Filtro regex inválido — considerando todos os envelopes.")
        ja = self._carregar_historico() if cfg["ignorar_baixados"] else set()

        self.tab_env.limpar()
        self.envelopes.clear()
        self.iid_por_chave.clear()
        for i, d in enumerate(linhas):
            if d["chave"] in self.envelopes:
                continue
            if filtro and not filtro.search(d["nome"]):
                status = "Fora do filtro"
            elif d["chave"] in ja:
                status = "Já baixado"
            else:
                status = "Pendente"
            iid = f"e{i}"
            self.envelopes[d["chave"]] = d
            self.iid_por_chave[d["chave"]] = iid
            self.tab_env.inserir(iid, (status, d["nome"], d["colaborador"], d["para"], d["data"],
                                       d["hora"], "", d["guid"]), status)
        self.tab_env.ordenar("data", forcar_asc=True)
        self._atualizar_kpis()

    def _sair(self):
        if self.worker:
            self.worker.parar.set()
            self.worker.enviar("sair")
            self.worker.join(timeout=5)
        self.destroy()


if __name__ == "__main__":
    App().mainloop()

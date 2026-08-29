import os
import re
import ssl
import ctypes
import mimetypes
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from urllib.parse import urlparse, parse_qs, unquote
from urllib.request import Request, urlopen
from html import unescape
from pathlib import Path

APP_TITLE = "Baixador de Arquivos - Outlook SafeLinks"

HTML_LINK_RE = re.compile(
    r'<a\b[^>]*?href=["\']([^"\']+)["\'][^>]*>(.*?)</a>',
    re.IGNORECASE | re.DOTALL
)
MARKDOWN_LINK_RE = re.compile(r'\[([^\]]+)\]\((https?://[^)]+)\)', re.IGNORECASE)
URL_RE = re.compile(r'https?://[^\s<>"\']+', re.IGNORECASE)
RTF_HYPERLINK_RE = re.compile(r'HYPERLINK\s+(?:\\\\")?["\']?(https?://[^"\'\s}\\]+)', re.IGNORECASE)
RTF_FIELD_RE = re.compile(r'\\fldinst\s*\{?[^{}]*?HYPERLINK\s+["\']?(https?://[^"\'\s}\\]+)', re.IGNORECASE | re.DOTALL)
TAG_RE = re.compile(r'<[^>]+>')

CONTENT_TYPE_EXT = {
    "application/pdf": ".pdf",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.ms-excel.sheet.binary.macroenabled.12": ".xlsb",
    "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsm",
    "application/zip": ".zip",
    "application/xml": ".xml",
    "text/xml": ".xml",
    "text/csv": ".csv",
    "text/plain": ".txt",
    "text/html": ".html",
    "application/octet-stream": "",
}


def _read_registered_clipboard_format(format_name):
    if os.name != "nt":
        return None

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    fmt = user32.RegisterClipboardFormatW(format_name)
    if not fmt or not user32.OpenClipboard(None):
        return None

    try:
        if not user32.IsClipboardFormatAvailable(fmt):
            return None
        handle = user32.GetClipboardData(fmt)
        if not handle:
            return None

        kernel32.GlobalLock.restype = ctypes.c_void_p
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return None
        try:
            size = kernel32.GlobalSize(handle)
            raw = ctypes.string_at(ptr, size)
        finally:
            kernel32.GlobalUnlock(handle)

        for encoding in ("utf-8", "utf-16-le", "cp1252", "latin1"):
            try:
                return raw.decode(encoding).replace("\x00", "")
            except Exception:
                pass
        return raw.decode("latin1", errors="replace").replace("\x00", "")
    finally:
        user32.CloseClipboard()


def get_clipboard_html():
    return _read_registered_clipboard_format("HTML Format")


def get_clipboard_rtf():
    for name in ("Rich Text Format", "Rich Text Format Without Objects", "text/rtf"):
        value = _read_registered_clipboard_format(name)
        if value:
            return value
    return None


def list_clipboard_formats():
    if os.name != "nt":
        return []

    user32 = ctypes.windll.user32
    formats = []
    if not user32.OpenClipboard(None):
        return formats

    try:
        fmt = 0
        while True:
            fmt = user32.EnumClipboardFormats(fmt)
            if fmt == 0:
                break
            if fmt == 1:
                name = "CF_TEXT"
            elif fmt == 13:
                name = "CF_UNICODETEXT"
            else:
                buf = ctypes.create_unicode_buffer(256)
                if user32.GetClipboardFormatNameW(fmt, buf, 256):
                    name = buf.value
                else:
                    name = f"Formato #{fmt}"
            formats.append(name)
    finally:
        user32.CloseClipboard()
    return formats


def get_open_outlook_email_html():
    if os.name != "nt":
        return None, "Disponível apenas no Windows."

    ps = r'''
$ErrorActionPreference = "Stop"
$outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$inspector = $outlook.ActiveInspector()
if ($null -eq $inspector) {
    throw "Nenhum e-mail está aberto em uma janela do Outlook."
}
$item = $inspector.CurrentItem
$subject = $item.Subject
$html = $item.HTMLBody
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Output "___SUBJECT___"
Write-Output $subject
Write-Output "___HTML___"
Write-Output $html
'''

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )
        if result.returncode != 0:
            err = (result.stderr or result.stdout or "").strip()
            return None, err or "Não foi possível acessar o Outlook."

        output = result.stdout
        if "___HTML___" not in output:
            return None, "O Outlook não retornou o corpo HTML."

        before, html = output.split("___HTML___", 1)
        subject = ""
        if "___SUBJECT___" in before:
            subject = before.split("___SUBJECT___", 1)[1].strip()
        return html.strip(), subject
    except subprocess.TimeoutExpired:
        return None, "Tempo excedido ao tentar acessar o Outlook."
    except Exception as e:
        return None, str(e)


def normalize_url(url):
    if not url:
        return ""
    url = unescape(url.strip())
    url = url.replace(r"\&", "&").replace(r"\/", "/").replace("\\'", "'")
    return url.rstrip(").,;]}>\\")


def decode_safelink(url):
    url = normalize_url(url)
    try:
        parsed = urlparse(url)
    except Exception:
        return url

    if "safelinks.protection.outlook.com" in parsed.netloc.lower():
        target = parse_qs(parsed.query).get("url", [None])[0]
        if target:
            return unquote(target)
    return url


def clean_html_label(label_html):
    label = TAG_RE.sub(" ", label_html)
    label = unescape(label)
    return re.sub(r"\s+", " ", label).strip() or "arquivo"


def deduplicate(items):
    result, seen = [], set()
    for item in items:
        target = normalize_url(item["target_url"])
        if not target or target in seen:
            continue
        seen.add(target)
        result.append(item)
    return result


def extract_links_from_html(html, only_download=True):
    items = []
    for href, label_html in HTML_LINK_RE.findall(html or ""):
        href = normalize_url(href)
        if not href.lower().startswith(("http://", "https://")):
            continue
        label = clean_html_label(label_html)
        target = decode_safelink(href)
        if only_download:
            low = (label + " " + href + " " + target).lower()
            if "download" not in low and "portal.gi.inf.br" not in low:
                continue
        items.append({"label": label, "safe_url": href, "target_url": target})
    return deduplicate(items)


def extract_links_from_rtf(rtf, only_download=True):
    urls = []
    for regex in (RTF_FIELD_RE, RTF_HYPERLINK_RE):
        urls.extend(regex.findall(rtf or ""))
    if not urls:
        urls = URL_RE.findall(rtf or "")

    items = []
    for i, raw in enumerate(urls, start=1):
        url = normalize_url(raw)
        target = decode_safelink(url)
        low = (url + " " + target).lower()
        if only_download and "safelinks.protection.outlook.com" not in low and "portal.gi.inf.br" not in low:
            continue
        items.append({"label": f"Download {i}", "safe_url": url, "target_url": target})
    return deduplicate(items)


def extract_links_from_text(text, only_download=True):
    items = []
    markdown = MARKDOWN_LINK_RE.findall(text or "")
    if markdown:
        for label, raw in markdown:
            url = normalize_url(raw)
            target = decode_safelink(url)
            low = (label + " " + url + " " + target).lower()
            if only_download and "download" not in low and "portal.gi.inf.br" not in low and "safelinks.protection.outlook.com" not in low:
                continue
            items.append({"label": label.strip() or "arquivo", "safe_url": url, "target_url": target})
    else:
        for i, raw in enumerate(URL_RE.findall(text or ""), start=1):
            url = normalize_url(raw)
            target = decode_safelink(url)
            low = (url + " " + target).lower()
            if only_download and "portal.gi.inf.br" not in low and "safelinks.protection.outlook.com" not in low:
                continue
            items.append({"label": f"Download {i}", "safe_url": url, "target_url": target})
    return deduplicate(items)


def sanitize_filename(name):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name or "").strip()
    name = re.sub(r"\s+", " ", name).rstrip(". ")
    return name or "arquivo"


def filename_from_content_disposition(cd):
    if not cd:
        return None
    m = re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", cd, re.IGNORECASE)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    m = re.search(r'filename\s*=\s*"([^"]+)"', cd, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"filename\s*=\s*([^;]+)", cd, re.IGNORECASE)
    if m:
        return m.group(1).strip().strip('"')
    return None


def guess_extension(content_type, url):
    ctype = (content_type or "").split(";")[0].strip().lower()
    if ctype in CONTENT_TYPE_EXT:
        return CONTENT_TYPE_EXT[ctype]
    ext = mimetypes.guess_extension(ctype) if ctype else None
    if ext:
        return ext
    path_ext = os.path.splitext(urlparse(url).path)[1]
    return path_ext if path_ext and len(path_ext) <= 8 else ""


def unique_path(folder, filename):
    candidate = folder / filename
    if not candidate.exists():
        return candidate
    stem, suffix, i = candidate.stem, candidate.suffix, 2
    while True:
        alt = folder / f"{stem}_{i}{suffix}"
        if not alt.exists():
            return alt
        i += 1


def download_file(item, folder, log):
    target_url = item["target_url"]
    label = item["label"]
    log(f"Baixando: {label}")

    req = Request(
        target_url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/151.0 Safari/537.36",
            "Accept": "*/*",
        },
        method="GET",
    )
    context = ssl.create_default_context()

    with urlopen(req, timeout=90, context=context) as response:
        content = response.read()
        content_type = response.headers.get("Content-Type", "")
        cd = response.headers.get("Content-Disposition", "")
        header_name = filename_from_content_disposition(cd)

        if header_name:
            filename = sanitize_filename(header_name)
        else:
            base = re.sub(r"(?i)\bdownload\b", "", label).strip(" -_")
            filename = sanitize_filename(base)
            ext = guess_extension(content_type, target_url)
            if ext and not filename.lower().endswith(ext.lower()):
                filename += ext

        if not os.path.splitext(filename)[1]:
            ext = guess_extension(content_type, target_url)
            if ext:
                filename += ext

        save_path = unique_path(folder, filename)
        save_path.write_bytes(content)

        if "text/html" in (content_type or "").lower():
            log(f"AVISO: {save_path.name} veio como HTML; pode ser página intermediária/login.")
        else:
            log(f"OK: {save_path.name} ({len(content):,} bytes)")
        return save_path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1080x800")
        self.minsize(920, 680)
        self.folder_var = tk.StringVar(value=str(Path.home() / "Downloads"))
        self.only_download_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Use 'Capturar do Clipboard' ou abra o e-mail e use 'Ler e-mail aberto'.")
        self.current_items = []
        self.current_source = ""
        self._build_ui()

    def _build_ui(self):
        main = ttk.Frame(self, padding=12)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="Baixador de arquivos do Outlook", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        ttk.Label(main, text="Tenta HTML, RTF, texto e também pode ler diretamente o e-mail aberto no Outlook clássico.").pack(anchor="w", pady=(2, 10))

        toolbar = ttk.Frame(main)
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Button(toolbar, text="Capturar do Clipboard", command=self.capture_clipboard).pack(side="left")
        ttk.Button(toolbar, text="Ler e-mail aberto no Outlook", command=self.read_open_outlook_email).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Diagnosticar Clipboard", command=self.diagnose_clipboard).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Limpar", command=self.clear_all).pack(side="left", padx=(6, 0))
        ttk.Checkbutton(toolbar, text="Somente links de download/NFE", variable=self.only_download_var).pack(side="left", padx=(16, 0))

        content_frame = ttk.LabelFrame(main, text="Conteúdo / links identificados")
        content_frame.pack(fill="both", expand=True, pady=(0, 10))
        self.text = tk.Text(content_frame, height=16, wrap="word", font=("Consolas", 9))
        self.text.pack(fill="both", expand=True, padx=8, pady=8)

        folder_frame = ttk.Frame(main)
        folder_frame.pack(fill="x")
        ttk.Label(folder_frame, text="Pasta de destino:").pack(side="left")
        ttk.Entry(folder_frame, textvariable=self.folder_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(folder_frame, text="Escolher pasta", command=self.choose_folder).pack(side="left")

        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(action_frame, text="Baixar todos", command=self.start_download).pack(side="left")
        ttk.Label(action_frame, textvariable=self.status_var).pack(side="left", padx=(14, 0))

        ttk.Separator(main).pack(fill="x", pady=10)
        ttk.Label(main, text="Log:").pack(anchor="w")
        self.log_box = tk.Text(main, height=15, wrap="word", state="disabled", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=False)

    def log(self, msg):
        def write():
            self.log_box.config(state="normal")
            self.log_box.insert("end", str(msg) + "\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(0, write)

    def clear_log(self):
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")

    def show_items(self, items, source):
        self.current_items = items
        self.current_source = source
        self.text.delete("1.0", "end")
        self.clear_log()

        if not items:
            self.status_var.set("Nenhum link de arquivo foi localizado.")
            self.text.insert("1.0", f"Origem analisada: {source}\n\nNenhum link de arquivo foi localizado.")
            return

        lines = [f"Origem: {source}", f"Links únicos encontrados: {len(items)}", ""]
        for i, item in enumerate(items, start=1):
            lines += [f"{i:02d}. {item['label']}", f"    {item['target_url']}", ""]
        self.text.insert("1.0", "\n".join(lines))
        self.status_var.set(f"{len(items)} arquivo(s) único(s) encontrado(s).")

    def capture_clipboard(self):
        only_download = self.only_download_var.get()

        html = get_clipboard_html()
        if html:
            items = extract_links_from_html(html, only_download)
            if items:
                self.show_items(items, "Clipboard HTML")
                return

        rtf = get_clipboard_rtf()
        if rtf:
            items = extract_links_from_rtf(rtf, only_download)
            if items:
                self.show_items(items, "Clipboard RTF")
                return

        try:
            text = self.clipboard_get()
        except tk.TclError:
            text = ""
        if text:
            items = extract_links_from_text(text, only_download)
            if items:
                self.show_items(items, "Clipboard texto")
                return

        formats = list_clipboard_formats()
        self.show_items([], "Clipboard")
        msg = "Não encontrei hyperlinks utilizáveis no conteúdo copiado.\n\nAbra o e-mail em uma janela separada e use 'Ler e-mail aberto no Outlook'."
        if formats:
            msg += "\n\nFormatos encontrados:\n" + "\n".join(formats[:12])
        messagebox.showwarning(APP_TITLE, msg)

    def diagnose_clipboard(self):
        formats = list_clipboard_formats()
        self.clear_log()
        self.log("Formatos disponíveis no clipboard:")
        if formats:
            for fmt in formats:
                self.log(f" - {fmt}")
        else:
            self.log(" - Nenhum formato identificado")
        self.log("")
        self.log(f"HTML Format: {'SIM' if get_clipboard_html() else 'NÃO'}")
        self.log(f"RTF: {'SIM' if get_clipboard_rtf() else 'NÃO'}")
        try:
            txt = self.clipboard_get()
            self.log(f"Texto: SIM ({len(txt)} caracteres)")
        except Exception:
            self.log("Texto: NÃO")
        self.status_var.set("Diagnóstico exibido no log.")

    def read_open_outlook_email(self):
        self.status_var.set("Lendo e-mail aberto no Outlook...")
        self.update_idletasks()
        html, info = get_open_outlook_email_html()
        if not html:
            messagebox.showerror(APP_TITLE, "Não consegui ler o e-mail aberto.\n\n" + str(info) + "\n\nEssa opção funciona com o Outlook clássico para Windows.")
            self.status_var.set("Não foi possível acessar o Outlook.")
            return

        items = extract_links_from_html(html, self.only_download_var.get())
        self.show_items(items, f"Outlook aberto — {info or '(sem assunto)'}")
        if not items:
            messagebox.showwarning(APP_TITLE, "Consegui ler o e-mail, mas não localizei links compatíveis.")

    def choose_folder(self):
        folder = filedialog.askdirectory(initialdir=self.folder_var.get())
        if folder:
            self.folder_var.set(folder)

    def clear_all(self):
        self.current_items = []
        self.current_source = ""
        self.text.delete("1.0", "end")
        self.clear_log()
        self.status_var.set("Use 'Capturar do Clipboard' ou 'Ler e-mail aberto no Outlook'.")

    def start_download(self):
        if not self.current_items:
            messagebox.showwarning(APP_TITLE, "Primeiro capture os links ou leia o e-mail aberto.")
            return
        folder = Path(self.folder_var.get()).expanduser()
        folder.mkdir(parents=True, exist_ok=True)
        items = list(self.current_items)
        self.status_var.set(f"Baixando {len(items)} arquivo(s)...")
        threading.Thread(target=self._download_worker, args=(items, folder), daemon=True).start()

    def _download_worker(self, items, folder):
        ok = errors = 0
        self.log("")
        self.log("=" * 72)
        self.log(f"Origem: {self.current_source}")
        self.log(f"Arquivos: {len(items)}")
        self.log(f"Destino: {folder}")
        self.log("=" * 72)

        for idx, item in enumerate(items, start=1):
            try:
                self.log("")
                self.log(f"[{idx}/{len(items)}]")
                download_file(item, folder, self.log)
                ok += 1
            except Exception as e:
                errors += 1
                self.log(f"ERRO: {item['label']} -> {e}")

        self.after(0, lambda: self.status_var.set(f"Finalizado: {ok} sucesso(s), {errors} erro(s)."))
        if errors == 0:
            self.after(0, lambda: messagebox.showinfo(APP_TITLE, f"Download concluído.\n\nArquivos: {ok}\nPasta: {folder}"))
        else:
            self.after(0, lambda: messagebox.showwarning(APP_TITLE, f"Finalizado com erros.\n\nSucessos: {ok}\nErros: {errors}\n\nVeja o log."))


if __name__ == "__main__":
    App().mainloop()

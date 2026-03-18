import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import queue
import subprocess
import shutil
import re
import urllib.request
import io
import ssl
import webbrowser

try:
    import yt_dlp
except ImportError:
    messagebox.showerror("Erreur", "yt-dlp n'est pas installé.\nLancez: pip install yt-dlp")
    sys.exit(1)

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


PLATFORMS      = ["YouTube", "TikTok", "Instagram", "Twitter", "Facebook",
                  "Vimeo", "Dailymotion", "Telegram"]
PRORES_PROFILE = "1"
PRORES_EXT     = ".mov"

BG      = "#0f0f13"
CARD    = "#1a1a24"
CARD2   = "#13131b"
BORDER  = "#2a2a3a"
SEL     = "#2e2416"       # fond carte sélectionnée
ACCENT  = "#c8922a"
ACCENT2 = "#e8b84b"
TEXT    = "#e8e8f0"
MUTED   = "#5a5a7a"
GREEN   = "#4ade80"


def find_ffmpeg():
    ff = shutil.which("ffmpeg")
    if ff:
        return ff
    for p in ["/opt/homebrew/bin/ffmpeg", "/usr/local/bin/ffmpeg", "/usr/bin/ffmpeg"]:
        if os.path.isfile(p):
            return p
    return None


def fmt_dur(secs):
    if not secs:
        return ""
    secs = int(secs)
    h, m, s = secs // 3600, (secs % 3600) // 60, secs % 60
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def fmt_views(n):
    if not n:
        return ""
    if n >= 1_000_000:
        return f"{n/1_000_000:.1f}M vues"
    if n >= 1_000:
        return f"{n/1_000:.0f}K vues"
    return f"{n} vues"


def fmt_size(b):
    if not b:
        return ""
    if b >= 1_000_000_000:
        return f"{b/1e9:.1f} GB"
    return f"{b/1_000_000:.0f} MB"


# ── Modèle vidéo ─────────────────────────────────────────────────────────────

class VideoInfo:
    def __init__(self, url, raw):
        self.url      = url
        self.title    = raw.get("title", url[:60])
        self.dur      = fmt_dur(raw.get("duration"))
        self.views    = fmt_views(raw.get("view_count"))
        self.channel  = raw.get("channel") or raw.get("uploader", "")
        self.thumb_url = raw.get("thumbnail", "")
        self.formats  = self._parse(raw.get("formats", []))
        self.sel      = 0   # index sélectionné

    def _parse(self, raw):
        out = []
        for h in [2160, 1440, 1080, 720, 480, 360, 240, 144]:
            vids = [f for f in raw
                    if f.get("height") == h
                    and f.get("vcodec", "none") not in ("none", None, "")]
            if vids:
                best = max(vids, key=lambda f: f.get("tbr") or 0)
                sz = best.get("filesize") or best.get("filesize_approx")
                out.append({
                    "label": f"{h}p", "badge": "VIDÉO",
                    "spec": f"bestvideo[height<={h}]+bestaudio/best[height<={h}]",
                    "size": fmt_size(sz), "audio": False,
                })
        audio = [f for f in raw
                 if f.get("vcodec", "none") in ("none", None, "")
                 and f.get("acodec", "none") not in ("none", None, "")]
        if audio:
            out.append({"label": "Audio MP3", "badge": "AUDIO",
                        "spec": "bestaudio/best", "size": "", "audio": True})
        if not out:
            out.append({"label": "Meilleur", "badge": "VIDÉO",
                        "spec": "bestvideo+bestaudio/best", "size": "", "audio": False})
        return out


# ── App ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gandalf OSINT")
        self.geometry("1100x1100")
        self.resizable(True, True)
        self.configure(bg=BG)

        self.output_dir    = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        self.prores_var    = tk.BooleanVar(value=True)   # kept for compat
        # "none" | "prores" | "mp4"
        self.transcode_var = tk.StringVar(value="prores")
        self.xlsx_rows: list[dict] = []
        self.log_q       = queue.Queue()
        self.running     = False
        self.analysing   = False
        self.ffmpeg_path = find_ffmpeg()
        self._last_file  = None
        self._video_infos: list[VideoInfo] = []
        self._card_btns   = {}   # card_idx → list of (border_f, inner_f, all_labels)

        self._gif_frames = []
        self._gif_delays = []
        self._gif_idx    = 0

        self._build()
        self._poll()

        if not self.ffmpeg_path:
            self._log("⚠  ffmpeg introuvable — brew install ffmpeg")
            for rb in self._tc_radios:
                rb.configure(state="disabled")
            self.transcode_var.set("none")

    # ── Construction ─────────────────────────────────────────────────────────

    def _build(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("G.TCheckbutton", background=CARD,
                        foreground=MUTED, font=("Helvetica", 11))
        style.map("G.TCheckbutton", background=[("active", CARD)],
                  foreground=[("active", ACCENT)])
        style.configure("G.Horizontal.TProgressbar",
                        troughcolor=BORDER, background=ACCENT, thickness=4)
        style.configure("TScrollbar", background=CARD,
                        troughcolor=BG, arrowcolor=MUTED)

        # ── Header ──
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", pady=(30, 0))

        if HAS_PIL:
            try:
                gif = Image.open(resource_path("gandalf.gif"))
                for i in range(gif.n_frames):
                    gif.seek(i)
                    self._gif_frames.append(ImageTk.PhotoImage(gif.copy().convert("RGBA")))
                    self._gif_delays.append(gif.info.get("duration", 60))
                self._gif_lbl = tk.Label(hdr, bg=BG)
                self._gif_lbl.pack()
                self._tick_gif()
            except Exception:
                pass

        tk.Label(hdr, text="Gandalf OSINT", bg=BG, fg=ACCENT2,
                 font=("Georgia", 38, "bold")).pack(pady=(12, 0))
        tk.Label(hdr, text="VIDÉOS  •  AUDIO  •  OSINT",
                 bg=BG, fg=MUTED, font=("Helvetica", 11)).pack(pady=(4, 0))

        pills = tk.Frame(hdr, bg=BG)
        pills.pack(pady=(12, 0))
        for p in PLATFORMS:
            tk.Label(pills, text=p, bg=CARD, fg=MUTED,
                     font=("Helvetica", 9), padx=11, pady=4).pack(side="left", padx=4)

        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=40, pady=(20, 0))

        # ── Zone URL ──
        input_card = tk.Frame(self, bg=CARD, padx=28, pady=22)
        input_card.pack(padx=40, fill="x", pady=(18, 0))

        self._ph = "Colle un ou plusieurs liens (un par ligne)..."
        self.url_box = tk.Text(input_card, bg=CARD2, fg=MUTED,
                               insertbackground=ACCENT,
                               font=("Helvetica", 12), relief="flat",
                               height=4, wrap="none", padx=12, pady=10,
                               highlightthickness=1,
                               highlightbackground=BORDER,
                               highlightcolor=ACCENT)
        self.url_box.insert("1.0", self._ph)
        self.url_box.pack(fill="x")
        self.url_box.bind("<FocusIn>",  self._clear_ph)
        self.url_box.bind("<FocusOut>", self._add_ph)

        self.analyse_btn = tk.Button(input_card, text="✦  Analyser",
                                     bg=ACCENT, fg="#0a0a0e",
                                     font=("Helvetica", 13, "bold"),
                                     relief="flat", padx=24, pady=10,
                                     cursor="hand2",
                                     activebackground=ACCENT2,
                                     activeforeground="#0a0a0e",
                                     command=self._analyse)
        self.analyse_btn.pack(anchor="w", pady=(16, 0))

        # ── Carte XLSX ──
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=40, pady=(18, 0))

        xlsx_card = tk.Frame(self, bg=CARD, padx=28, pady=22)
        xlsx_card.pack(padx=40, fill="x", pady=(14, 0))

        head_row = tk.Frame(xlsx_card, bg=CARD)
        head_row.pack(fill="x")
        tk.Label(head_row, text="📊  IMPORT EXCEL (.xlsx)",
                 bg=CARD, fg=ACCENT2, font=("Helvetica", 12, "bold")).pack(side="left")

        tk.Label(xlsx_card,
                 text="Colonnes requises : key  ·  publication_date  ·  description  ·  location  ·  link",
                 bg=CARD, fg=MUTED, font=("Helvetica", 9)).pack(anchor="w", pady=(4, 12))

        btn_row = tk.Frame(xlsx_card, bg=CARD)
        btn_row.pack(fill="x")
        tk.Button(btn_row, text="📂  Importer .xlsx",
                  bg=CARD2, fg=ACCENT, font=("Helvetica", 11, "bold"),
                  relief="flat", padx=16, pady=8,
                  highlightthickness=1, highlightbackground=BORDER,
                  cursor="hand2", command=self._import_xlsx).pack(side="left")

        self._xlsx_lbl = tk.Label(btn_row,
                                  text="Aucun fichier importé",
                                  bg=CARD, fg=MUTED, font=("Helvetica", 10))
        self._xlsx_lbl.pack(side="left", padx=(16, 0))

        self._xlsx_dl_btn = tk.Button(xlsx_card,
                                      text="✦  Télécharger (Excel)",
                                      bg=ACCENT, fg="#0a0a0e",
                                      font=("Georgia", 13, "bold"),
                                      relief="flat", padx=24, pady=10,
                                      cursor="hand2",
                                      activebackground=ACCENT2,
                                      activeforeground="#0a0a0e",
                                      command=self._download_xlsx)
        self._xlsx_dl_btn.pack(anchor="w", pady=(14, 0))

        # ── Résultats scrollables ──
        self.results_outer = tk.Frame(self, bg=BG)
        self.results_outer.pack(padx=40, fill="both", expand=True, pady=(14, 0))

        self.canvas = tk.Canvas(self.results_outer, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(self.results_outer, orient="vertical",
                            command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.results_frame = tk.Frame(self.canvas, bg=BG)
        self._cwin = self.canvas.create_window((0, 0), window=self.results_frame, anchor="nw")
        self.results_frame.bind("<Configure>",
            lambda _: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
            lambda e: self.canvas.itemconfig(self._cwin, width=e.width))
        self.canvas.bind_all("<MouseWheel>",
            lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # ── Barre du bas ──
        bot = tk.Frame(self, bg=CARD, padx=28, pady=16)
        bot.pack(padx=40, fill="x", pady=(12, 0))

        left = tk.Frame(bot, bg=CARD)
        left.pack(side="left", fill="y")

        style.configure("G.TRadiobutton", background=CARD,
                        foreground=MUTED, font=("Helvetica", 11))
        style.map("G.TRadiobutton", background=[("active", CARD)],
                  foreground=[("active", ACCENT)])

        tk.Label(left, text="TRANSCODAGE", bg=CARD, fg=MUTED,
                 font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(0, 6))

        self._tc_radios = []
        for val, label in [
            ("none",   "Aucun — garder le fichier source"),
            ("prores", "Apple ProRes 422 LT  (.mov)"),
            ("mp4",    "MP4 H.264  — compatible Premiere Pro  (.mp4)"),
        ]:
            rb = ttk.Radiobutton(left, text=label,
                                 variable=self.transcode_var, value=val,
                                 style="G.TRadiobutton")
            rb.pack(anchor="w", pady=2)
            self._tc_radios.append(rb)

        folder_row = tk.Frame(left, bg=CARD)
        folder_row.pack(anchor="w", pady=(8, 0))
        tk.Label(folder_row, text="📁", bg=CARD, fg=MUTED,
                 font=("Helvetica", 20)).pack(side="left")
        tk.Label(folder_row, textvariable=self.output_dir,
                 bg=CARD, fg=MUTED, font=("Helvetica", 16)).pack(side="left", padx=(10, 14))
        tk.Button(folder_row, text="Changer",
                  bg=BG, fg=ACCENT, font=("Helvetica", 14),
                  relief="flat", padx=18, pady=6,
                  highlightthickness=0, bd=0, cursor="hand2",
                  command=self._choose_dir).pack(side="left")

        self.dl_btn = tk.Button(bot, text="✦  Télécharger",
                                bg=ACCENT, fg="#0a0a0e",
                                font=("Georgia", 14, "bold"),
                                relief="flat", padx=36, pady=14,
                                cursor="hand2",
                                activebackground=ACCENT2,
                                activeforeground="#0a0a0e",
                                command=self._download_all)
        self.dl_btn.pack(side="right")

        # ── Progression ──
        prog = tk.Frame(self, bg=BG)
        prog.pack(padx=40, fill="x", pady=(10, 0))
        self.progress = ttk.Progressbar(prog, style="G.Horizontal.TProgressbar",
                                        mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        self.status_lbl = tk.Label(prog, text="", bg=BG, fg=MUTED,
                                   font=("Helvetica", 9), anchor="w")
        self.status_lbl.pack(anchor="w", pady=(4, 0))

        # ── Log ──
        self.log_box = tk.Text(self, bg=CARD, fg=GREEN,
                               font=("Courier", 8), relief="flat",
                               height=4, state="disabled",
                               wrap="word", padx=10, pady=8,
                               highlightthickness=1,
                               highlightbackground=BORDER)
        self.log_box.pack(padx=40, fill="x", pady=(10, 0))

        # ── Footer ──
        footer = tk.Frame(self, bg=BG)
        footer.pack(pady=(10, 12))
        tk.Label(footer, text="Fait par ", bg=BG, fg="#2a2a3a",
                 font=("Helvetica", 9)).pack(side="left")
        lnk = tk.Label(footer, text="Théotime Massard", bg=BG, fg=MUTED,
                       font=("Helvetica", 9, "underline"), cursor="hand2")
        lnk.pack(side="left")
        lnk.bind("<Enter>", lambda _: lnk.configure(fg=ACCENT))
        lnk.bind("<Leave>", lambda _: lnk.configure(fg=MUTED))
        lnk.bind("<Button-1>", lambda _: webbrowser.open(
            "https://youtube.com/@theotimemassard?si=FFIVqWw1d5wMWeXq"))

    # ── GIF ───────────────────────────────────────────────────────────────────

    def _tick_gif(self):
        if not self._gif_frames:
            return
        self._gif_lbl.configure(image=self._gif_frames[self._gif_idx])
        delay = self._gif_delays[self._gif_idx]
        self._gif_idx = (self._gif_idx + 1) % len(self._gif_frames)
        self.after(delay, self._tick_gif)

    # ── Placeholder ───────────────────────────────────────────────────────────

    def _clear_ph(self, _):
        if self.url_box.get("1.0", "end").strip() == self._ph:
            self.url_box.delete("1.0", "end")
            self.url_box.configure(fg=TEXT)

    def _add_ph(self, _):
        if not self.url_box.get("1.0", "end").strip():
            self.url_box.insert("1.0", self._ph)
            self.url_box.configure(fg=MUTED)

    # ── Analyse ───────────────────────────────────────────────────────────────

    def _analyse(self):
        if self.analysing or self.running:
            return
        raw  = self.url_box.get("1.0", "end").strip()
        urls = [u.strip() for u in raw.splitlines()
                if u.strip() and u.strip() != self._ph]
        if not urls:
            messagebox.showwarning("Aucun lien", "Colle au moins un lien.")
            return

        # Reset résultats
        for w in self.results_frame.winfo_children():
            w.destroy()
        self._video_infos.clear()
        self._card_btns.clear()

        self.analysing = True
        self.analyse_btn.configure(state="disabled", bg=BORDER)
        self._set_status("Analyse en cours…")
        threading.Thread(target=self._fetch_all, args=(urls,), daemon=True).start()

    def _fetch_all(self, urls):
        ydl_opts = {"quiet": True, "no_warnings": True, "skip_download": True}
        if self.ffmpeg_path:
            ydl_opts["ffmpeg_location"] = os.path.dirname(self.ffmpeg_path)

        for i, url in enumerate(urls):
            self._set_status(f"Analyse [{i+1}/{len(urls)}] {url[:60]}")
            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    raw = ydl.extract_info(url, download=False)
                info = VideoInfo(url, raw)
                self._video_infos.append(info)
                self.after(0, self._add_video_card, len(self._video_infos)-1, info)
            except Exception as e:
                self._log(f"Erreur analyse : {e}")

        self.analysing = False
        self.analyse_btn.configure(state="normal", bg=ACCENT)
        self._set_status(f"{len(self._video_infos)} vidéo(s) prête(s) — choisissez la qualité.")

    # ── Carte vidéo ───────────────────────────────────────────────────────────

    def _add_video_card(self, idx: int, info: VideoInfo):
        card = tk.Frame(self.results_frame, bg=CARD, padx=22, pady=18)
        card.pack(fill="x", pady=(0, 10))

        # ── Header : miniature + infos ──
        top = tk.Frame(card, bg=CARD)
        top.pack(fill="x")

        # Miniature
        thumb_lbl = tk.Label(top, bg="#0a0a0e", width=18, height=6)
        thumb_lbl.pack(side="left", padx=(0, 18))

        if HAS_PIL and info.thumb_url:
            threading.Thread(target=self._fetch_thumb,
                             args=(info.thumb_url, thumb_lbl), daemon=True).start()

        # Infos
        inf = tk.Frame(top, bg=CARD)
        inf.pack(side="left", fill="x", expand=True, anchor="n")

        tk.Label(inf, text=info.title, bg=CARD, fg=TEXT,
                 font=("Helvetica", 13, "bold"),
                 anchor="w", wraplength=660, justify="left").pack(anchor="w")

        pills = tk.Frame(inf, bg=CARD)
        pills.pack(anchor="w", pady=(8, 0))
        for txt in [info.dur, info.views, info.channel]:
            if txt:
                tk.Label(pills, text=txt, bg=BORDER, fg=MUTED,
                         font=("Helvetica", 9), padx=8, pady=3).pack(side="left", padx=(0, 6))

        # ── Sélection qualité ──
        tk.Label(card, text="CHOISIR LA QUALITÉ", bg=CARD, fg=MUTED,
                 font=("Helvetica", 9, "bold")).pack(anchor="w", pady=(18, 8))

        grid = tk.Frame(card, bg=CARD)
        grid.pack(anchor="w", fill="x")

        self._card_btns[idx] = []
        self._build_quality_grid(grid, idx, info)

    def _build_quality_grid(self, parent, card_idx, info: VideoInfo):
        MAX_ROW = 7
        row_f = None
        btns  = self._card_btns[card_idx]

        for i, fmt in enumerate(info.formats):
            if i % MAX_ROW == 0:
                row_f = tk.Frame(parent, bg=CARD)
                row_f.pack(anchor="w", pady=(0, 6))

            # Bordure externe (change couleur si sélectionné)
            brd = tk.Frame(row_f, bg=BORDER, padx=1, pady=1)
            brd.pack(side="left", padx=(0, 8))

            # Contenu interne
            inn_bg = CARD2
            inn = tk.Frame(brd, bg=inn_bg, padx=16, pady=10, cursor="hand2")
            inn.pack()

            badge_col = GREEN if fmt["audio"] else ACCENT
            lbl_badge = tk.Label(inn, text=fmt["badge"], bg=inn_bg,
                                 fg=badge_col, font=("Helvetica", 7, "bold"))
            lbl_badge.pack()

            lbl_res = tk.Label(inn, text=fmt["label"], bg=inn_bg,
                               fg=TEXT, font=("Helvetica", 14, "bold"))
            lbl_res.pack()

            widgets = [inn, lbl_badge, lbl_res]

            if fmt["size"]:
                lbl_sz = tk.Label(inn, text=fmt["size"], bg=inn_bg,
                                  fg=MUTED, font=("Helvetica", 8))
                lbl_sz.pack()
                widgets.append(lbl_sz)

            btns.append((brd, inn, widgets))

            def _click(event, ci=card_idx, fi=i):
                self._select_quality(ci, fi)

            for w in widgets:
                w.bind("<Button-1>", _click)

        # Sélection par défaut
        self._select_quality(card_idx, 0)

    def _select_quality(self, card_idx, fmt_idx):
        info = self._video_infos[card_idx]
        info.sel = fmt_idx
        btns = self._card_btns.get(card_idx, [])
        for i, (brd, inn, widgets) in enumerate(btns):
            selected = (i == fmt_idx)
            brd.configure(bg=ACCENT if selected else BORDER)
            inn.configure(bg=SEL if selected else CARD2)
            for w in widgets:
                w.configure(bg=SEL if selected else CARD2)

    # ── Miniature ─────────────────────────────────────────────────────────────

    def _fetch_thumb(self, url, label):
        try:
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            data = urllib.request.urlopen(req, timeout=6, context=ctx).read()
            img  = Image.open(io.BytesIO(data)).convert("RGB")
            img  = img.resize((130, 74), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            label.after(0, lambda: (
                label.configure(image=photo, bg="#0a0a0e", width=130, height=74),
                setattr(label, "_img", photo)
            ))
        except Exception:
            pass

    # ── Téléchargement ────────────────────────────────────────────────────────

    def _download_all(self):
        if self.running or not self._video_infos:
            if not self._video_infos:
                messagebox.showinfo("Aucune vidéo", "Lance d'abord l'analyse.")
            return
        self.running = True
        self.dl_btn.configure(state="disabled", bg=BORDER)
        self.analyse_btn.configure(state="disabled")
        threading.Thread(target=self._batch, daemon=True).start()

    def _batch(self):
        total = len(self._video_infos)
        for n, info in enumerate(self._video_infos, 1):
            self._set_status(f"[{n}/{total}] Téléchargement…")
            self._run(n, info, total)
        self._set_status(f"Terminé — {total} fichier(s).")
        self.progress["value"] = 100
        self.running = False
        self.dl_btn.configure(state="normal", bg=ACCENT)
        self.analyse_btn.configure(state="normal", bg=ACCENT)

    def _run(self, n, info: VideoInfo, total):
        self._last_file = None
        fmt = info.formats[info.sel]
        try:
            opts = self._ydl_opts(fmt["spec"], n)
            self._log(f"[{n}/{total}] {info.title[:60]}")
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([info.url])

            dl = self._last_file
            if not dl or not os.path.isfile(dl):
                raise FileNotFoundError("Fichier introuvable.")

            tc = self.transcode_var.get()
            if tc != "none" and not fmt["audio"] and self._has_video(dl):
                self._transcode(n, dl, total, tc)
            else:
                self._log(f"[{n}/{total}] ✓ {os.path.basename(dl)}")

        except yt_dlp.utils.DownloadError as e:
            self._log(f"[{n}/{total}] ✗ {e}")
        except Exception as e:
            self._log(f"[{n}/{total}] ✗ {e}")

    # ── yt-dlp ────────────────────────────────────────────────────────────────

    def _ydl_opts(self, spec, n):
        out = os.path.join(self.output_dir.get(), "%(title)s.%(ext)s")
        pp  = []
        if spec == "bestaudio/best":
            pp = [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "192"}]

        opts = {
            "format": spec, "outtmpl": out, "postprocessors": pp,
            "progress_hooks":      [lambda d, _n=n: self._dl_hook(d, _n)],
            "postprocessor_hooks": [self._pp_hook],
            "quiet": True, "no_warnings": False, "merge_output_format": "mp4",
        }
        if self.ffmpeg_path:
            opts["ffmpeg_location"] = os.path.dirname(self.ffmpeg_path)
        return opts

    def _dl_hook(self, d, n):
        if d["status"] == "downloading":
            try:
                pct = float(d.get("_percent_str", "0%").strip().replace("%", ""))
                self.progress["value"] = pct * 0.7
            except ValueError:
                pass
            self._set_status(
                f"[{n}] {d.get('_percent_str','').strip()}  "
                f"{d.get('_speed_str','?')}  ETA {d.get('_eta_str','?')}"
            )

    def _pp_hook(self, d):
        if d["status"] == "finished":
            fp = (d.get("info_dict") or {}).get("filepath") or d.get("filename")
            if fp:
                self._last_file = fp

    # ── ProRes ────────────────────────────────────────────────────────────────

    def _has_video(self, path):
        ffprobe = shutil.which("ffprobe") or self.ffmpeg_path.replace("ffmpeg", "ffprobe")
        if not (ffprobe and os.path.isfile(ffprobe)):
            return True
        try:
            out = subprocess.check_output(
                [ffprobe, "-v", "error", "-select_streams", "v:0",
                 "-show_entries", "stream=codec_type",
                 "-of", "default=noprint_wrappers=1:nokey=1", path],
                stderr=subprocess.DEVNULL, text=True)
            return "video" in out
        except Exception:
            return True

    def _get_duration(self, path):
        ffprobe = shutil.which("ffprobe") or self.ffmpeg_path.replace("ffmpeg", "ffprobe")
        if not (ffprobe and os.path.isfile(ffprobe)):
            return None
        try:
            out = subprocess.check_output(
                [ffprobe, "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", path],
                stderr=subprocess.DEVNULL, text=True)
            return float(out.strip())
        except Exception:
            return None

    def _transcode(self, n, src, total, mode="prores", final_name=None):
        folder = os.path.dirname(src)
        base   = os.path.splitext(src)[0]

        if mode == "prores":
            if final_name:
                dst = os.path.join(folder, final_name + PRORES_EXT)
            else:
                dst = (base + "_prores" + PRORES_EXT) if src.endswith(PRORES_EXT) else (base + PRORES_EXT)
            self._log(f"[{n}/{total}] ProRes 422 LT → {os.path.basename(dst)}")
            cmd = [self.ffmpeg_path, "-y", "-i", src,
                   "-c:v", "prores_ks", "-profile:v", PRORES_PROFILE,
                   "-vendor", "apl0", "-bits_per_mb", "8000", "-pix_fmt", "yuv422p10le",
                   "-c:a", "pcm_s24le", dst]
        else:  # mp4 h264 premiere
            if final_name:
                dst = os.path.join(folder, final_name + ".mp4")
            else:
                dst = base + "_premiere.mp4" if src.endswith(".mp4") else base + ".mp4"
            self._log(f"[{n}/{total}] MP4 H.264 Premiere → {os.path.basename(dst)}")
            cmd = [self.ffmpeg_path, "-y", "-i", src,
                   "-c:v", "libx264", "-preset", "slow", "-crf", "18",
                   "-pix_fmt", "yuv420p",
                   "-c:a", "aac", "-b:a", "320k",
                   "-movflags", "+faststart",
                   dst]

        # Si src == dst (excel mode : on réencode vers le même nom) → temp file
        if os.path.abspath(src) == os.path.abspath(dst):
            tmp = src + ".tmp_tc" + os.path.splitext(dst)[1]
            cmd[-1] = tmp
        else:
            tmp = None

        duration = self._get_duration(src)
        try:
            proc = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                                    universal_newlines=True, bufsize=1)
            t_re = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")
            for line in proc.stderr:
                m = t_re.search(line)
                if m and duration:
                    h, mn, s, cs = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                    pct = min((h*3600+mn*60+s+cs/100)/duration*100, 99)
                    self.progress["value"] = 70 + pct * 0.3
                    label = "ProRes 422 LT" if mode == "prores" else "MP4 H.264"
                    self._set_status(f"[{n}/{total}] {label}… {pct:.0f}%")
            proc.wait()
            if proc.returncode == 0:
                lbl = "ProRes 422 LT" if mode == "prores" else "MP4 Premiere"
                # Si on a utilisé un fichier temp, on le renomme
                if tmp and os.path.isfile(tmp):
                    if os.path.isfile(dst):
                        os.remove(dst)
                    os.rename(tmp, dst)
                try:
                    if os.path.abspath(src) != os.path.abspath(dst):
                        os.remove(src)
                except OSError:
                    pass
                self._log(f"[{n}/{total}] ✓ {lbl} : {os.path.basename(dst)}")
            else:
                self._log(f"[{n}/{total}] ✗ ffmpeg erreur {proc.returncode}")
        except FileNotFoundError:
            self._log("ffmpeg introuvable.")

    # ── XLSX ──────────────────────────────────────────────────────────────────

    def _import_xlsx(self):
        path = filedialog.askopenfilename(
            title="Importer un fichier Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")])
        if not path:
            return
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("Erreur", "openpyxl manquant.\npip install openpyxl")
            return
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active

            # Cherche la ligne d'en-tête (contient "key" ET "link")
            header_row = None
            for i, row in enumerate(ws.iter_rows(max_row=15, values_only=True), 1):
                vals = [str(v).strip().lower() if v else "" for v in row]
                if "key" in vals and "link" in vals:
                    header_row = i
                    headers = [str(v).strip() if v else "" for v in row]
                    break

            if not header_row:
                messagebox.showerror("Erreur",
                    "Impossible de trouver les en-têtes.\n"
                    "Vérifiez que le fichier contient les colonnes : key, link, publication_date, description, location")
                return

            required = {"key", "link", "publication_date", "description", "location"}
            low_headers = [h.lower() for h in headers]
            missing = required - set(low_headers)
            if missing:
                messagebox.showerror("Colonnes manquantes",
                    f"Introuvables : {', '.join(sorted(missing))}\n"
                    f"Trouvées : {', '.join(h for h in headers if h)}")
                return

            # Normalise les headers en minuscules pour l'accès
            norm = [h.lower() for h in headers]

            self.xlsx_rows = []
            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                d = {norm[i]: row[i] for i in range(len(norm)) if i < len(row)}
                if d.get("link") and str(d["link"]).startswith("http"):
                    self.xlsx_rows.append(d)

            fname = os.path.basename(path)
            self._xlsx_lbl.configure(
                text=f"✓  {fname}  —  {len(self.xlsx_rows)} lien(s) trouvé(s)",
                fg=GREEN)
            self._log(f"Excel importé : {fname} ({len(self.xlsx_rows)} lignes)")

        except Exception as e:
            messagebox.showerror("Erreur lecture xlsx", str(e))

    def _download_xlsx(self):
        if not self.xlsx_rows:
            messagebox.showinfo("Aucune donnée", "Importe d'abord un fichier Excel.")
            return
        if self.running:
            return
        self.running = True
        self._xlsx_dl_btn.configure(state="disabled", bg=BORDER)
        self.dl_btn.configure(state="disabled")
        threading.Thread(target=self._batch_xlsx, daemon=True).start()

    def _batch_xlsx(self):
        total = len(self.xlsx_rows)
        for n, row in enumerate(self.xlsx_rows, 1):
            url  = str(row.get("link", "")).strip()
            name = self._build_name(row)
            self._set_status(f"[{n}/{total}] {name[:70]}")
            self._run_named(n, url, name, total)
        self._set_status(f"Terminé — {total} fichier(s) (Excel).")
        self.progress["value"] = 100
        self.running = False
        self._xlsx_dl_btn.configure(state="normal", bg=ACCENT)
        self.dl_btn.configure(state="normal", bg=ACCENT)

    def _build_name(self, row) -> str:
        """Construit le nom de fichier : key_publication_date_description_location"""
        def s(v):
            if v is None:
                return ""
            v = str(v).strip()
            # Seul "/" est interdit dans les noms de fichiers macOS
            return v.replace("/", "-")

        key   = s(row.get("key"))
        date  = s(row.get("publication_date"))
        desc  = s(row.get("description"))
        loc   = s(row.get("location"))
        return f"{key}_{date}_{desc}_{loc}"

    def _run_named(self, n, url, custom_name, total):
        """Télécharge un lien avec un nom de fichier personnalisé."""
        self._last_file = None
        try:
            out  = os.path.join(self.output_dir.get(), f"{custom_name}.%(ext)s")
            opts = {
                "format": "bestvideo+bestaudio/best",
                "outtmpl": out,
                "postprocessors": [],
                "progress_hooks":      [lambda d, _n=n: self._dl_hook(d, _n)],
                "postprocessor_hooks": [self._pp_hook],
                "quiet": True, "no_warnings": False,
                "merge_output_format": "mp4",
                "restrictfilenames": False,
            }
            if self.ffmpeg_path:
                opts["ffmpeg_location"] = os.path.dirname(self.ffmpeg_path)

            self._log(f"[{n}/{total}] {url[:70]}")
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([url])

            dl = self._last_file
            if not dl or not os.path.isfile(dl):
                raise FileNotFoundError("Fichier introuvable après téléchargement.")

            tc = self.transcode_var.get()
            if tc != "none" and self._has_video(dl):
                self._transcode(n, dl, total, tc, final_name=custom_name)
            else:
                self._log(f"[{n}/{total}] ✓  {os.path.basename(dl)}")

        except yt_dlp.utils.DownloadError as e:
            self._log(f"[{n}/{total}] ✗  {e}")
        except Exception as e:
            self._log(f"[{n}/{total}] ✗  {e}")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _choose_dir(self):
        d = filedialog.askdirectory(initialdir=self.output_dir.get())
        if d:
            self.output_dir.set(d)

    def _set_status(self, msg):
        self.status_lbl.configure(text=msg)

    def _log(self, msg):
        self.log_q.put(msg)

    def _poll(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log_box.configure(state="normal")
                self.log_box.insert("end", msg + "\n")
                self.log_box.see("end")
                self.log_box.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(150, self._poll)


if __name__ == "__main__":
    App().mainloop()

import threading
import json
import os
import sys
import subprocess
import shutil
import re

import webview

try:
    import yt_dlp
except ImportError:
    pass

PRORES_PROFILE = "1"
PRORES_EXT     = ".mov"

# On Windows, hide console windows spawned by subprocess
SUBPROCESS_FLAGS = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def find_ffmpeg():
    # On Windows the bundled binary is ffmpeg.exe
    for name in ("ffmpeg.exe", "ffmpeg"):
        bundled = resource_path(name)
        if os.path.isfile(bundled):
            if sys.platform != "win32":
                os.chmod(bundled, 0o755)
            return bundled
    ff = shutil.which("ffmpeg")
    if ff:
        return ff
    if IS_WINDOWS:
        for p in [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "ffmpeg", "bin", "ffmpeg.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""), "ffmpeg", "bin", "ffmpeg.exe"),
            r"C:\ffmpeg\bin\ffmpeg.exe",
        ]:
            if os.path.isfile(p):
                return p
    else:
        for p in ["/opt/homebrew/bin/ffmpeg", "/usr/local/bin/ffmpeg", "/usr/bin/ffmpeg"]:
            if os.path.isfile(p):
                return p
    return None


def find_ffprobe(ffmpeg_path):
    """Derive ffprobe path from ffmpeg path, handling .exe on Windows."""
    if not ffmpeg_path:
        return None
    d = os.path.dirname(ffmpeg_path)
    probe = "ffprobe.exe" if IS_WINDOWS else "ffprobe"
    candidate = os.path.join(d, probe)
    if os.path.isfile(candidate):
        return candidate
    return shutil.which("ffprobe")


def fmt_dur(secs):
    if not secs: return ""
    secs = int(secs)
    h, m, s = secs // 3600, (secs % 3600) // 60, secs % 60
    return f"{h}:{m:02d}:{s:02d}" if h else f"{m}:{s:02d}"


def fmt_views(n):
    if not n: return ""
    if n >= 1_000_000: return f"{n/1_000_000:.1f}M vues"
    if n >= 1_000:     return f"{n/1_000:.0f}K vues"
    return f"{n} vues"


def fmt_size(b):
    if not b: return ""
    if b >= 1_000_000_000: return f"{b/1e9:.1f} GB"
    return f"{b/1_000_000:.0f} MB"


class VideoInfo:
    def __init__(self, url, raw):
        self.url       = url
        self.title     = raw.get("title", url[:60])
        self.dur       = fmt_dur(raw.get("duration"))
        self.views     = fmt_views(raw.get("view_count"))
        self.channel   = raw.get("channel") or raw.get("uploader", "")
        self.thumb_url = raw.get("thumbnail", "")
        self.formats   = self._parse(raw.get("formats", []))
        self.sel       = 0

    def _parse(self, raw):
        out = []
        for h in [2160, 1440, 1080, 720, 480, 360, 240, 144]:
            vids = [f for f in raw
                    if f.get("height") == h
                    and f.get("vcodec", "none") not in ("none", None, "")]
            if vids:
                best = max(vids, key=lambda f: f.get("tbr") or 0)
                sz = best.get("filesize") or best.get("filesize_approx")
                out.append({"label": f"{h}p", "badge": "VIDÉO",
                            "spec": f"bestvideo[height<={h}]+bestaudio/best[height<={h}]",
                            "size": fmt_size(sz), "audio": False})
        if not out:
            any_vid = [f for f in raw if f.get("vcodec", "none") not in ("none", None, "")]
            if any_vid:
                out.append({"label": "Meilleure", "badge": "VIDÉO",
                            "spec": "bestvideo+bestaudio/best", "size": "", "audio": False})
        audio = [f for f in raw
                 if f.get("vcodec", "none") in ("none", None, "")
                 and f.get("acodec", "none") not in ("none", None, "")]
        if audio:
            out.append({"label": "Audio MP3", "badge": "AUDIO",
                        "spec": "bestaudio/best", "size": "", "audio": True})
        if not out:
            out.append({"label": "Meilleure", "badge": "VIDÉO",
                        "spec": "bestvideo+bestaudio/best", "size": "", "audio": False})
        return out

    def to_dict(self):
        return {"url": self.url, "title": self.title, "dur": self.dur,
                "views": self.views, "channel": self.channel,
                "thumb_url": self.thumb_url, "formats": self.formats, "sel": self.sel}


class Api:
    def __init__(self):
        self._window        = None
        self.output_dir     = os.path.expanduser("~/Downloads")
        self.ffmpeg_path    = find_ffmpeg()
        self._video_infos   = []
        self._xlsx_rows     = []
        self.running        = False
        self.analysing      = False
        self._last_file     = None
        self.transcode_mode = "prores"
        self.xlsx_quality   = "bestvideo+bestaudio/best"

    # ── JS → Python ──────────────────────────────────────────────────────────

    def get_initial_state(self):
        return {
            "output_dir":       self.output_dir,
            "ffmpeg_available": bool(self.ffmpeg_path),
            "transcode_mode":   self.transcode_mode,
        }

    def analyse_urls(self, urls_text):
        if self.analysing or self.running:
            return
        urls = [u.strip() for u in urls_text.splitlines() if u.strip()]
        if not urls:
            return
        self._video_infos.clear()
        self.analysing = True
        threading.Thread(target=self._fetch_all, args=(urls,), daemon=True).start()

    def set_quality(self, card_idx, fmt_idx):
        card_idx, fmt_idx = int(card_idx), int(fmt_idx)
        if 0 <= card_idx < len(self._video_infos):
            self._video_infos[card_idx].sel = fmt_idx

    def set_transcode(self, mode):
        self.transcode_mode = mode

    def set_xlsx_quality(self, spec):
        self.xlsx_quality = spec

    def download_all(self):
        if self.running or not self._video_infos:
            return
        self.running = True
        self._emit("download_started")
        threading.Thread(target=self._batch, daemon=True).start()

    def import_xlsx(self):
        result = self._window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=('Excel files (*.xlsx *.xls)', 'All files (*.*)')
        )
        if not result:
            return None
        path = result[0]
        try:
            import openpyxl
        except ImportError:
            return {"error": "openpyxl manquant — pip install openpyxl"}
        try:
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            header_row = None
            for i, row in enumerate(ws.iter_rows(max_row=15, values_only=True), 1):
                vals = [str(v).strip().lower() if v else "" for v in row]
                if "key" in vals and "link" in vals:
                    header_row = i
                    headers = [str(v).strip() if v else "" for v in row]
                    break
            if not header_row:
                return {"error": "En-têtes introuvables (key + link requis)"}
            norm = [h.lower() for h in headers]
            self._xlsx_rows = []
            for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
                d = {norm[i]: row[i] for i in range(len(norm)) if i < len(row)}
                if d.get("link") and str(d["link"]).startswith("http"):
                    self._xlsx_rows.append(d)
            return {"filename": os.path.basename(path), "count": len(self._xlsx_rows)}
        except Exception as e:
            return {"error": str(e)}

    def download_xlsx(self):
        if not self._xlsx_rows or self.running:
            return
        self.running = True
        self._emit("download_started")
        threading.Thread(target=self._batch_xlsx, daemon=True).start()

    def choose_dir(self):
        result = self._window.create_file_dialog(
            webview.FOLDER_DIALOG, directory=self.output_dir)
        if result:
            self.output_dir = result[0]
            return result[0]
        return ""

    def open_link(self, url):
        if sys.platform == "win32":
            os.startfile(url)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", url])
        else:
            subprocess.Popen(["xdg-open", url])

    # ── Internal ─────────────────────────────────────────────────────────────

    def _emit(self, event, data=None):
        payload = json.dumps(data) if data is not None else "null"
        if self._window:
            self._window.evaluate_js(f"window.onEvent('{event}', {payload})")

    def _fetch_all(self, urls):
        ydl_opts = {"quiet": True, "no_warnings": True, "skip_download": True}
        if self.ffmpeg_path:
            ydl_opts["ffmpeg_location"] = os.path.dirname(self.ffmpeg_path)
        for i, url in enumerate(urls):
            self._emit("status", f"Analyse [{i+1}/{len(urls)}] {url[:60]}")
            try:
                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    raw = ydl.extract_info(url, download=False)
                info = VideoInfo(url, raw)
                self._video_infos.append(info)
                self._emit("video_analysed", info.to_dict())
            except Exception as e:
                self._emit("log", f"Erreur analyse : {e}")
        self.analysing = False
        self._emit("analyse_done", len(self._video_infos))
        self._emit("status", f"{len(self._video_infos)} vidéo(s) prête(s).")

    def _batch(self):
        total = len(self._video_infos)
        for n, info in enumerate(self._video_infos, 1):
            self._emit("status", f"[{n}/{total}] Téléchargement…")
            self._run(n, info, total)
        self._emit("status", f"Terminé — {total} fichier(s).")
        self._emit("progress", 100)
        self.running = False
        self._emit("download_done")

    def _run(self, n, info, total):
        self._last_file = None
        fmt = info.formats[info.sel]
        try:
            opts = self._ydl_opts(fmt["spec"], n)
            self._emit("log", f"[{n}/{total}] {info.title[:60]}")
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([info.url])
            dl = self._last_file
            if not dl or not os.path.isfile(dl):
                raise FileNotFoundError("Fichier introuvable.")
            tc = self.transcode_mode
            if tc != "none" and not fmt["audio"] and self._has_video(dl):
                self._transcode(n, dl, total, tc)
            else:
                self._emit("log", f"[{n}/{total}] ✓ {os.path.basename(dl)}")
        except Exception as e:
            self._emit("log", f"[{n}/{total}] ✗ {e}")

    def _ydl_opts(self, spec, n):
        out = os.path.join(self.output_dir, "%(title)s.%(ext)s")
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
                self._emit("progress", pct * 0.7)
            except ValueError:
                pass
            self._emit("status",
                f"[{n}] {d.get('_percent_str','').strip()}  "
                f"{d.get('_speed_str','?')}  ETA {d.get('_eta_str','?')}")

    def _pp_hook(self, d):
        if d["status"] == "finished":
            fp = (d.get("info_dict") or {}).get("filepath") or d.get("filename")
            if fp:
                self._last_file = fp

    def _has_video(self, path):
        ffprobe = find_ffprobe(self.ffmpeg_path)
        if not ffprobe:
            return True
        try:
            out = subprocess.check_output(
                [ffprobe, "-v", "error", "-select_streams", "v:0",
                 "-show_entries", "stream=codec_type",
                 "-of", "default=noprint_wrappers=1:nokey=1", path],
                stderr=subprocess.DEVNULL, text=True,
                creationflags=SUBPROCESS_FLAGS)
            return "video" in out
        except Exception:
            return True

    def _get_duration(self, path):
        ffprobe = find_ffprobe(self.ffmpeg_path)
        if not ffprobe:
            return None
        try:
            out = subprocess.check_output(
                [ffprobe, "-v", "error", "-show_entries", "format=duration",
                 "-of", "default=noprint_wrappers=1:nokey=1", path],
                stderr=subprocess.DEVNULL, text=True,
                creationflags=SUBPROCESS_FLAGS)
            return float(out.strip())
        except Exception:
            return None

    def _transcode(self, n, src, total, mode="prores", final_name=None):
        folder = os.path.dirname(src)
        base   = os.path.splitext(src)[0]
        if mode == "prores":
            dst = (base + "_prores" + PRORES_EXT) if src.endswith(PRORES_EXT) else (base + PRORES_EXT)
            if final_name:
                dst = os.path.join(folder, final_name + PRORES_EXT)
            self._emit("log", f"[{n}/{total}] ProRes 422 LT → {os.path.basename(dst)}")
            cmd = [self.ffmpeg_path, "-y", "-i", src,
                   "-c:v", "prores_ks", "-profile:v", PRORES_PROFILE,
                   "-vendor", "apl0", "-bits_per_mb", "8000", "-pix_fmt", "yuv422p10le",
                   "-c:a", "pcm_s24le", dst]
        else:
            dst = base + "_premiere.mp4" if src.endswith(".mp4") else base + ".mp4"
            if final_name:
                dst = os.path.join(folder, final_name + ".mp4")
            self._emit("log", f"[{n}/{total}] MP4 H.264 → {os.path.basename(dst)}")
            cmd = [self.ffmpeg_path, "-y", "-i", src,
                   "-c:v", "libx264", "-preset", "slow", "-crf", "18",
                   "-pix_fmt", "yuv420p", "-c:a", "aac", "-b:a", "320k",
                   "-movflags", "+faststart", dst]

        if os.path.abspath(src) == os.path.abspath(dst):
            tmp = src + ".tmp_tc" + os.path.splitext(dst)[1]
            cmd[-1] = tmp
        else:
            tmp = None

        duration = self._get_duration(src)
        try:
            proc = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                                    universal_newlines=True, bufsize=1,
                                    creationflags=SUBPROCESS_FLAGS)
            t_re = re.compile(r"time=(\d+):(\d+):(\d+)\.(\d+)")
            for line in proc.stderr:
                m = t_re.search(line)
                if m and duration:
                    h, mn, s, cs = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
                    pct = min((h*3600+mn*60+s+cs/100)/duration*100, 99)
                    self._emit("progress", 70 + pct * 0.3)
            proc.wait()
            if proc.returncode == 0:
                if tmp and os.path.isfile(tmp):
                    if os.path.isfile(dst): os.remove(dst)
                    os.rename(tmp, dst)
                try:
                    if os.path.abspath(src) != os.path.abspath(dst):
                        os.remove(src)
                except OSError:
                    pass
                lbl = "ProRes 422 LT" if mode == "prores" else "MP4 Premiere"
                self._emit("log", f"[{n}/{total}] ✓ {lbl} : {os.path.basename(dst)}")
            else:
                self._emit("log", f"[{n}/{total}] ✗ ffmpeg erreur {proc.returncode}")
        except FileNotFoundError:
            self._emit("log", "ffmpeg introuvable.")

    def _batch_xlsx(self):
        total = len(self._xlsx_rows)
        for n, row in enumerate(self._xlsx_rows, 1):
            url  = str(row.get("link", "")).strip()
            name = self._build_name(row)
            self._emit("status", f"[{n}/{total}] {name[:70]}")
            self._run_named(n, url, name, total)
        self._emit("status", f"Terminé — {total} fichier(s) (Excel).")
        self._emit("progress", 100)
        self.running = False
        self._emit("download_done")

    def _build_name(self, row):
        def s(v):
            if v is None: return ""
            return str(v).strip().replace("/", "-")
        return (f"{s(row.get('key'))}_{s(row.get('publication_date'))}_"
                f"{s(row.get('description'))}_{s(row.get('location'))}")

    def _run_named(self, n, url, custom_name, total):
        self._last_file = None
        is_audio = self.xlsx_quality == "bestaudio/best"
        try:
            out  = os.path.join(self.output_dir, f"{custom_name}.%(ext)s")
            pp   = [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3",
                     "preferredquality": "192"}] if is_audio else []
            opts = {
                "format": self.xlsx_quality, "outtmpl": out, "postprocessors": pp,
                "progress_hooks":      [lambda d, _n=n: self._dl_hook(d, _n)],
                "postprocessor_hooks": [self._pp_hook],
                "quiet": True, "no_warnings": False,
                "merge_output_format": "mp4", "restrictfilenames": False,
            }
            if self.ffmpeg_path:
                opts["ffmpeg_location"] = os.path.dirname(self.ffmpeg_path)
            self._emit("log", f"[{n}/{total}] {url[:70]}")
            with yt_dlp.YoutubeDL(opts) as ydl:
                ydl.download([url])
            dl = self._last_file
            if not dl or not os.path.isfile(dl):
                raise FileNotFoundError("Fichier introuvable.")
            tc = self.transcode_mode
            if tc != "none" and not is_audio and self._has_video(dl):
                self._transcode(n, dl, total, tc, final_name=custom_name)
            else:
                self._emit("log", f"[{n}/{total}] ✓  {os.path.basename(dl)}")
        except Exception as e:
            self._emit("log", f"[{n}/{total}] ✗  {e}")


if __name__ == "__main__":
    api = Api()
    html_path = resource_path("app.html")
    icon_path = resource_path("icon.icns")

    # Convert to a proper file:// URI (handles Windows drive letters & backslashes)
    from pathlib import Path
    html_uri = Path(html_path).as_uri()

    window = webview.create_window(
        "Gandalf OSINT",
        html_path,
        js_api=api,
        width=960,
        height=880,
        min_size=(700, 600),
        background_color="#0a0a0d",
    )
    api._window = window
    webview.start()

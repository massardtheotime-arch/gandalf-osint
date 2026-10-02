"""Weekly health check: runs the app's real analyse+download code path on a
set of sample URLs (one per source) and reports pass/fail per source.

Usage: python3 tools/health_check.py
Exit code 0 = everything passed, 1 = at least one source failed.
"""
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import downloader_app as app  # noqa: E402

SAMPLES = [
    ("TikTok", "https://www.tiktok.com/@itamar_ben_gvir/video/7568955923486477576"),
    ("YouTube", "https://youtu.be/NFnTWw83ae8?si=40IfevzUEFd_qlwD"),
    ("X", "https://x.com/FabulasGuy/status/2092611204309229929"),
    ("Instagram post", "https://www.instagram.com/p/DciTNggDWZX"),
    ("Instagram reel", "https://www.instagram.com/reels/DcpT69wAGXH/"),
    ("Instagram carousel", "https://www.instagram.com/p/DbTUIqOjAWb/"),
]


def check(label, url):
    out = tempfile.mkdtemp(prefix="gandalf-health-")
    try:
        api = app.Api()
        api.output_dir = out
        api.transcode_mode = "none"
        api._fetch_all([url])
        if not api._video_infos:
            return False, "analyse: aucun résultat"
        total = len(api._video_infos)
        for n, info in enumerate(api._video_infos, 1):
            api._run(n, info, total)
        files = [f for f in os.listdir(out) if os.path.getsize(os.path.join(out, f)) > 0]
        if len(files) < total:
            return False, f"{len(files)} fichier(s) pour {total} élément(s)"
        return True, f"{len(files)} fichier(s)"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}"
    finally:
        shutil.rmtree(out, ignore_errors=True)


def main():
    import yt_dlp
    print(f"yt-dlp {yt_dlp.version.__version__}\n")
    failed = 0
    for label, url in SAMPLES:
        ok, detail = check(label, url)
        failed += not ok
        print(f"{'OK  ' if ok else 'FAIL'}  {label}: {detail}")
    print(f"\n{len(SAMPLES) - failed}/{len(SAMPLES)} OK")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())

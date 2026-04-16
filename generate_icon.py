"""Generate icon.ico for Windows from gollum.png."""
import os
import sys

try:
    from PIL import Image
except ImportError:
    print("Pillow non installe, icon.ico non genere.")
    sys.exit(0)

src = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gollum.png")
dst = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")

if os.path.isfile(dst):
    print(f"icon.ico existe deja.")
    sys.exit(0)

if not os.path.isfile(src):
    print(f"gollum.png introuvable, icon.ico non genere.")
    sys.exit(0)

img = Image.open(src).convert("RGBA")
sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
img.save(dst, format="ICO", sizes=sizes)
print(f"icon.ico genere depuis gollum.png")

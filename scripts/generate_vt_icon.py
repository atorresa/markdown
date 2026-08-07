"""Genera icon.ico recortando la V (de VIRTUAL) y la T (de TELCO) del logo original."""
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
LOGO_PATH = ROOT / "assets" / "logo.jpg"
ICON_PATH = ROOT / "icon.ico"

# Recortes (con un pequeño margen) detectados a partir del logo original 2048x387.
V_BOX = (110, 100, 288, 280)   # V de VIRTUAL (trazo contorno)
T_BOX = (1006, 108, 1172, 280)  # T de TELCO (trazo sólido); evita "L" (x<=1001) y "E" (x>=1175) vecinas


def make_transparent(im: Image.Image, white_cutoff=245, black_cutoff=205) -> Image.Image:
    """Convierte el fondo blanco en transparente con un borde suavizado."""
    im = im.convert("RGBA")
    pixels = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = pixels[x, y]
            gray = (r + g + b) / 3
            if gray >= white_cutoff:
                alpha = 0
            elif gray <= black_cutoff:
                alpha = 255
            else:
                alpha = int(255 * (white_cutoff - gray) / (white_cutoff - black_cutoff))
            pixels[x, y] = (r, g, b, alpha)
    return im


def main():
    """Recorta la V y la T del logo, las compone en un lienzo cuadrado y guarda icon.ico."""
    logo = Image.open(LOGO_PATH).convert("RGB")

    v_crop = make_transparent(logo.crop(V_BOX))
    t_crop = make_transparent(logo.crop(T_BOX))

    canvas_size = 256
    margin = 24
    gap = 20
    usable = canvas_size - 2 * margin

    # Una única escala (misma para ambas letras) calculada para que quepan una junto a otra
    # respetando el margen, sin deformar el grosor relativo de los trazos.
    raw_total_width = v_crop.width + gap + t_crop.width
    raw_max_height = max(v_crop.height, t_crop.height)
    scale = min(usable / raw_total_width, usable / raw_max_height)

    def scaled(im):
        return im.resize((round(im.width * scale), round(im.height * scale)), Image.LANCZOS)

    v_crop = scaled(v_crop)
    t_crop = scaled(t_crop)
    gap_scaled = round(gap * scale)

    total_width = v_crop.width + gap_scaled + t_crop.width
    canvas = Image.new("RGBA", (canvas_size, canvas_size), (0, 0, 0, 0))

    start_x = (canvas_size - total_width) // 2
    canvas.paste(v_crop, (start_x, (canvas_size - v_crop.height) // 2), v_crop)
    canvas.paste(
        t_crop,
        (start_x + v_crop.width + gap_scaled, (canvas_size - t_crop.height) // 2),
        t_crop,
    )

    canvas.save(ICON_PATH, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print(f"Icono generado en {ICON_PATH}")


if __name__ == "__main__":
    main()

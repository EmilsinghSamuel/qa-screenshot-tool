"""Generates icon.ico — called by the GitHub Actions build before PyInstaller."""
from PIL import Image, ImageDraw


def create():
    size = 256
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Windows-blue rounded-rectangle background (Windows 11 app icon style)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=52,
                           fill=(0, 120, 212, 255))

    sc = (255, 255, 255, 255)   # white scissors
    lw = 11                      # line / outline width
    rr = 30                      # finger-ring radius

    # Left finger ring
    lx, ly = 80, 88
    draw.ellipse([lx - rr, ly - rr, lx + rr, ly + rr], outline=sc, width=lw)

    # Right finger ring
    rx, ry = 176, 88
    draw.ellipse([rx - rr, ry - rr, rx + rr, ry + rr], outline=sc, width=lw)

    # Pivot (where blades cross)
    px, py, pr = 128, 168, 9
    draw.ellipse([px - pr, py - pr, px + pr, py + pr], fill=sc)

    # Left blade:  left-ring bottom → pivot → lower-right tip
    draw.line([(lx, ly + rr), (px, py)],   fill=sc, width=lw)
    draw.line([(px, py),      (208, 230)], fill=sc, width=lw)

    # Right blade: right-ring bottom → pivot → lower-left tip
    draw.line([(rx, ry + rr), (px, py)],   fill=sc, width=lw)
    draw.line([(px, py),      (48, 230)],  fill=sc, width=lw)

    # Multi-resolution ICO
    imgs = [img.resize((s, s), Image.LANCZOS) for s in (256, 64, 48, 32, 16)]
    imgs[0].save(
        "icon.ico",
        format="ICO",
        sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (256, 256)],
        append_images=imgs[1:],
    )
    print("icon.ico created.")


if __name__ == "__main__":
    create()

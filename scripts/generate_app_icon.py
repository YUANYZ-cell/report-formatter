"""Generate the Windows icon used by the local Report Formatter shortcut."""

from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "assets" / "report-formatter.ico"


def build_icon() -> Image.Image:
    image = Image.new("RGBA", (512, 512), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    # App tile
    draw.rounded_rectangle((56, 56, 456, 456), radius=104, fill="#232E5C")
    draw.rounded_rectangle((68, 68, 444, 444), radius=92, fill="#39478A")
    draw.rounded_rectangle((68, 68, 444, 214), radius=92, fill="#6173F4")
    draw.rectangle((68, 154, 444, 214), fill="#6173F4")

    # Paper with folded corner
    draw.rounded_rectangle((134, 104, 362, 392), radius=22, fill="#F9FAFF")
    draw.polygon(((306, 104), (362, 160), (306, 160)), fill="#DCE2FF")
    draw.line((306, 104, 306, 160, 362, 160), fill="#BFC8FA", width=8)
    draw.rounded_rectangle((168, 194, 318, 216), radius=11, fill="#7080EC")
    draw.rounded_rectangle((168, 239, 294, 255), radius=8, fill="#C6CEF7")
    draw.rounded_rectangle((168, 275, 276, 291), radius=8, fill="#C6CEF7")
    draw.rounded_rectangle((168, 311, 246, 327), radius=8, fill="#C6CEF7")

    # Validation mark
    draw.ellipse((278, 278, 386, 386), fill="#1FC6A0")
    draw.line((304, 332, 326, 354, 363, 309), fill="white", width=17, joint="curve")
    return image


if __name__ == "__main__":
    OUTPUT.parent.mkdir(exist_ok=True)
    build_icon().save(OUTPUT, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
    print(OUTPUT)

"""Genera una figura a partir de la salida real de pruebas y evaluación."""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
REPORTS_DIR = ROOT / "reports"
OUT = REPORTS_DIR / "figures" / "pruebas_automatizadas.png"
ANSI_ESCAPE = re.compile(r"\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])")


def load_font(size: int):
    """Carga una fuente monoespaciada disponible en Windows."""
    candidates = [
        Path("C:/Windows/Fonts/consola.ttf"),
        Path("C:/Windows/Fonts/lucon.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def run_command(label: str, command: list[str]) -> list[str]:
    """Ejecuta un comando y devuelve su salida; un fallo detiene la figura."""
    environment = os.environ.copy()
    environment.update(
        {
            "PYTHONDONTWRITEBYTECODE": "1",
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
        }
    )
    completed = subprocess.run(
        command,
        cwd=ROOT,
        env=environment,
        capture_output=True,
        check=False,
    )
    raw_output = b"\n".join(part for part in (completed.stdout, completed.stderr) if part)
    try:
        output = raw_output.decode("utf-8")
    except UnicodeDecodeError:
        output = raw_output.decode("cp1252", errors="replace")
    output = ANSI_ESCAPE.sub("", output).replace(str(ROOT), ".").strip()
    if completed.returncode != 0:
        raise RuntimeError(f"Falló {label}:\n{output}")

    lines = [f"> {label}"]
    for line in output.splitlines():
        wrapped = textwrap.wrap(line, width=112, replace_whitespace=False) or [""]
        lines.extend(wrapped)
    return lines


def render_output(lines: list[str]) -> None:
    """Renderiza la salida comprobada como figura legible para la memoria."""
    font = load_font(19)
    line_height = 34
    height = max(420, 90 + line_height * len(lines))
    image = Image.new("RGB", (1500, height), "#101419")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle(
        (30, 25, 1470, height - 25),
        radius=16,
        fill="#101419",
        outline="#3A4652",
        width=3,
    )
    y = 50
    for line in lines:
        lowered = line.lower()
        color = "#65D26E" if "passed" in lowered or "completada" in lowered else "#E7EDF3"
        if line.startswith(">"):
            color = "#6FC3DF"
        draw.text((65, y), line, font=font, fill=color)
        y += line_height

    OUT.parent.mkdir(parents=True, exist_ok=True)
    image.save(OUT)


def main() -> None:
    """Ejecuta verificaciones reales y genera la figura si todas pasan."""
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="pytest-evidence-", dir=REPORTS_DIR) as temp_dir:
        pytest_command = [
            sys.executable,
            "-m",
            "pytest",
            "-q",
            "-p",
            "no:cacheprovider",
            "--basetemp",
            str(Path(temp_dir) / "tmp"),
        ]
        lines = run_command("python -m pytest -q", pytest_command)

    lines.append("")
    lines.extend(
        run_command(
            "python scripts/evaluate.py",
            [sys.executable, str(ROOT / "scripts" / "evaluate.py")],
        )
    )
    render_output(lines)
    print(OUT)


if __name__ == "__main__":
    main()

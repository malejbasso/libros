#!/usr/bin/env python3
"""
convert.py — Descarga el .docx desde Google Drive y genera el index.html del libro.

Dependencias:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib python-docx beautifulsoup4 lxml

Variables de entorno requeridas (configúralas en GitHub Secrets o en .env local):
    GOOGLE_SERVICE_ACCOUNT_JSON  — JSON de la cuenta de servicio de Google (en una sola línea)
    DRIVE_FILE_ID                — ID del archivo .docx en Google Drive
                                   (está en la URL: drive.google.com/file/d/ESTE_ID/view)
"""

import json
import os
import re
import sys
import io
import html
from pathlib import Path

# ── Dependencias opcionales — instaladas en el entorno ──
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

try:
    from docx import Document
    from docx.oxml.ns import qn
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    from bs4 import BeautifulSoup
    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
BOOK_META = {
    "title":    os.environ.get("BOOK_TITLE",    "Actividades del Colegio"),
    "subtitle": os.environ.get("BOOK_SUBTITLE", "Registro de actividades y eventos"),
}

# Pages per "chapter" before forcing a new turn page (adjust to taste)
CHARS_PER_PAGE = 1800

# ──────────────────────────────────────────────
# 1. DESCARGA DESDE GOOGLE DRIVE
# ──────────────────────────────────────────────

def download_docx_from_drive(file_id: str, dest_path: str) -> None:
    """Descarga el .docx usando una cuenta de servicio de Google."""
    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client. Ejecuta: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib")

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise ValueError("Variable de entorno GOOGLE_SERVICE_ACCOUNT_JSON no encontrada.")

    sa_info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )

    service  = build("drive", "v3", credentials=creds, cache_discovery=False)
    request  = service.files().get_media(fileId=file_id)
    fh       = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    with open(dest_path, "wb") as f:
        f.write(fh.getvalue())

    print(f"✅ Archivo descargado: {dest_path}")


# ──────────────────────────────────────────────
# 2. CONVERSIÓN DOCX → PÁGINAS JSON
# ──────────────────────────────────────────────

def para_to_html(para) -> str:
    """Convierte un párrafo python-docx a fragmento HTML."""
    style  = para.style.name if para.style else ""
    text   = para.text.strip()

    if not text:
        return ""

    # Mapeo de estilos Word → tags HTML
    if style.startswith("Heading 1"):
        return f"<h1>{html.escape(text)}</h1>"
    if style.startswith("Heading 2"):
        return f"<h2>{html.escape(text)}</h2>"
    if style.startswith("Heading 3"):
        return f"<h3>{html.escape(text)}</h3>"
    if style in ("Quote", "Intense Quote"):
        return f"<blockquote>{html.escape(text)}</blockquote>"

    # Inline formatting (bold/italic/underline)
    parts = []
    for run in para.runs:
        t = html.escape(run.text)
        if run.bold:   t = f"<strong>{t}</strong>"
        if run.italic: t = f"<em>{t}</em>"
        if run.underline: t = f"<u>{t}</u>"
        parts.append(t)
    inner = "".join(parts) or html.escape(text)

    # List styles
    if "List" in style:
        return f"<li>{inner}</li>"

    return f"<p>{inner}</p>"


def table_to_html(table) -> str:
    rows_html = []
    for i, row in enumerate(table.rows):
        cells = []
        for cell in row.cells:
            tag = "th" if i == 0 else "td"
            cells.append(f"<{tag}>{html.escape(cell.text.strip())}</{tag}>")
        rows_html.append("<tr>" + "".join(cells) + "</tr>")
    return "<table>" + "".join(rows_html) + "</table>"


def docx_to_pages(docx_path: str) -> list:
    """
    Convierte el .docx en una lista de paginas para el libro.
    - NO genera portada: el contenido del Word se muestra tal cual.
    - Cada Heading 1 fuerza una nueva pagina (incluido en el contenido).
    - Una pagina en blanco al inicio alinea correctamente el spread de dos paginas.
    """
    if not HAS_DOCX:
        raise RuntimeError("Falta python-docx. Ejecuta: pip install python-docx")

    doc = Document(docx_path)
    pages = []

    # Pagina en blanco a la izquierda del primer spread,
    # asi la primera pagina de contenido queda a la derecha (como libro real)
    pages.append({"type": "blank"})

    buffer = []

    def flush_buffer():
        if not buffer:
            return
        combined = "".join(buffer)
        chunks = paginate_html(combined, CHARS_PER_PAGE)
        for chunk in chunks:
            pages.append({"type": "content", "content": chunk})
        buffer.clear()

    in_list = False

    for block in doc.element.body:
        tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

        if tag == "p":
            from docx.text.paragraph import Paragraph as DocxPara
            para = DocxPara(block, doc)
            style = para.style.name if para.style else ""
            text  = para.text.strip()

            if not text:
                continue

            # Heading 1 -> nueva pagina, heading incluido en el contenido
            if style.startswith("Heading 1"):
                if in_list:
                    buffer.append("</ul>")
                    in_list = False
                flush_buffer()
                buffer.append(f"<h1>{html.escape(text)}</h1>")
                continue

            fragment = para_to_html(para)
            if not fragment:
                continue

            if fragment.startswith("<li>"):
                if not in_list:
                    buffer.append("<ul>")
                    in_list = True
            else:
                if in_list:
                    buffer.append("</ul>")
                    in_list = False

            buffer.append(fragment)

        elif tag == "tbl":
            from docx.table import Table as DocxTable
            tbl = DocxTable(block, doc)
            if in_list:
                buffer.append("</ul>")
                in_list = False
            buffer.append(table_to_html(tbl))

    if in_list:
        buffer.append("</ul>")
    flush_buffer()

    # Contraportada final
    pages.append({"type": "back"})

    return pages


def paginate_html(html_content: str, max_chars: int) -> list:
    """Divide un bloque HTML largo en trozos de max_chars caracteres sin cortar tags."""
    # Dividir por párrafos/headings preservando el HTML
    tag_pattern = re.compile(r'(<(?:p|h[123]|blockquote|ul|table)[^>]*>.*?</(?:p|h[123]|blockquote|ul|table)>)', re.DOTALL)
    parts = tag_pattern.split(html_content)
    # Filtrar texto suelto
    elements = [p for p in parts if p.strip()]

    chunks  = []
    current = []
    count   = 0

    for el in elements:
        length = len(re.sub(r'<[^>]+>', '', el))  # solo contar texto visible
        if count + length > max_chars and current:
            chunks.append("".join(current))
            current = [el]
            count   = length
        else:
            current.append(el)
            count += length

    if current:
        chunks.append("".join(current))

    return chunks or [html_content]


# ──────────────────────────────────────────────
# 3. INYECCIÓN EN index.html
# ──────────────────────────────────────────────

def inject_pages_into_html(pages: list, template_path: str, output_path: str) -> None:
    """Reemplaza el placeholder __PAGES_JSON__ en el HTML con los datos reales."""
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    pages_json = json.dumps(pages, ensure_ascii=False, indent=2)
    output = template.replace("__PAGES_JSON__", pages_json)

    # Actualizar meta — reemplaza los valores en el objeto META del JS
    output = output.replace(
        'title:    "Actividades del Colegio"',
        f'title:    "{BOOK_META["title"]}"'
    )
    output = output.replace(
        'subtitle: "Registro de actividades y eventos"',
        f'subtitle: "{BOOK_META["subtitle"]}"'
    )
    # También actualizar el <title> del HTML
    output = output.replace(
        '<title>Libro Digital del Colegio</title>',
        f'<title>{BOOK_META["title"]}</title>'
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"✅ Libro generado: {output_path}  ({len(pages)} páginas)")


# ──────────────────────────────────────────────
# 4. MODO DEMO (sin Drive) — genera páginas de ejemplo
# ──────────────────────────────────────────────

def generate_demo_pages() -> list:
    """Genera páginas de ejemplo para ver el libro sin Drive/docx."""
    pages = [{"type": "cover"}]
    activities = [
        {
            "title": "Día del Libro",
            "date":  "12 de abril, 2025",
            "body":  "<p>Los alumnos de 5° básico organizaron una feria del libro en el patio central. Cada curso presentó su libro favorito con un afiche hecho a mano.</p><p>Participaron más de 120 estudiantes y se donaron 45 libros a la biblioteca del colegio.</p><blockquote>\"Leer es volar sin moverse del lugar.\"</blockquote>"
        },
        {
            "title": "Taller de Huerto Escolar",
            "date":  "3 de mayo, 2025",
            "body":  "<p>Con apoyo de padres voluntarios, los alumnos de 3° básico aprendieron a sembrar tomates, lechugas y hierbas aromáticas en el huerto del colegio.</p><p>Cada estudiante se llevó una pequeña maceta con semillas para continuar el aprendizaje en casa.</p>"
        },
        {
            "title": "Olimpiadas Matemáticas",
            "date":  "20 de mayo, 2025",
            "body":  "<p>Se realizó la segunda versión de las Olimpiadas Matemáticas internas. Participaron 8 equipos de distintos cursos.</p><table><tr><th>Lugar</th><th>Equipo</th><th>Puntos</th></tr><tr><td>1°</td><td>Los Calculadores</td><td>95</td></tr><tr><td>2°</td><td>Pi & Cía</td><td>87</td></tr><tr><td>3°</td><td>Suma y Sigue</td><td>81</td></tr></table>"
        },
        {
            "title": "Visita al Museo Nacional",
            "date":  "10 de junio, 2025",
            "body":  "<p>Los alumnos de 7° y 8° básico visitaron el Museo Nacional de Historia Natural en Santiago. El recorrido incluyó las salas de paleontología y biodiversidad.</p><p>Los estudiantes realizaron un taller de dibujo científico guiados por los educadores del museo.</p>"
        },
        {
            "title": "Festival de la Primavera",
            "date":  "21 de septiembre, 2025",
            "body":  "<p>El Festival de la Primavera reunió a todos los cursos en el patio para celebrar el inicio de la estación. Hubo presentaciones de danza, música y poesía.</p><p>El evento cerró con una suelta de globos elaborados por los propios estudiantes con materiales reciclados.</p>"
        },
    ]

    for i, act in enumerate(activities):
        pages.append({
            "type":           "content",
            "chapter":        act["title"],
            "chapter_label":  f"Actividad {i+1}  ·  {act['date']}",
            "content":        act["body"]
        })

    if len(pages) % 2 == 0:
        pages.append({"type": "back"})
    else:
        pages.append({"type": "back"})

    return pages


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    base_dir      = Path(__file__).parent
    template_path = base_dir / "index.html"
    output_path   = base_dir / "index.html"   # sobreescribe el template in-place
    tmp_docx      = base_dir / "tmp_document.docx"

    file_id = os.environ.get("DRIVE_FILE_ID")
    demo_mode = "--demo" in sys.argv or not file_id

    if demo_mode:
        print("ℹ️  Modo demo (sin Google Drive). Usa DRIVE_FILE_ID para conectar tu documento.")
        pages = generate_demo_pages()
    else:
        print(f"📥 Descargando documento desde Drive (ID: {file_id})…")
        download_docx_from_drive(file_id, str(tmp_docx))

        print("📖 Convirtiendo a páginas…")
        pages = docx_to_pages(str(tmp_docx))

        # Limpiar archivo temporal
        tmp_docx.unlink(missing_ok=True)

    inject_pages_into_html(pages, str(template_path), str(output_path))


if __name__ == "__main__":
    main()

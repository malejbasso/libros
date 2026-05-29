#!/usr/bin/env python3
"""
convert.py — Descarga el .docx desde Google Drive y genera el index.html del libro.

Dependencias:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib python-docx beautifulsoup4 lxml

Variables de entorno requeridas (configúralas en GitHub Secrets o en .env local):
    GOOGLE_SERVICE_ACCOUNT_JSON  — JSON de la cuenta de servicio de Google (en una sola línea)
    DRIVE_FILE_ID                — ID del archivo .docx en Google Drive
"""

import json, os, re, sys, io, html
from pathlib import Path

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# Chars of visible text per page before paginating
CHARS_PER_PAGE = 1800

# ──────────────────────────────────────────────
# 1. DESCARGA DESDE GOOGLE DRIVE
# ──────────────────────────────────────────────

def download_docx_from_drive(file_id: str, dest_path: str) -> None:
    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client.")
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise ValueError("Variable GOOGLE_SERVICE_ACCOUNT_JSON no encontrada.")
    sa_info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        sa_info, scopes=["https://www.googleapis.com/auth/drive.readonly"])
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    with open(dest_path, "wb") as f:
        f.write(fh.getvalue())
    print(f"✅ Archivo descargado: {dest_path}")

# ──────────────────────────────────────────────
# 2. EXTRACCIÓN DE METADATOS DEL WORD
# ──────────────────────────────────────────────

def extract_meta(doc) -> dict:
    """
    Extrae título y subtítulo del documento Word.
    - Título: primer párrafo con estilo Title, o primer Heading 1, o primer texto no vacío.
    - Subtítulo: primer párrafo con estilo Subtitle, o segundo párrafo no vacío si es corto.
    """
    title    = None
    subtitle = None

    for para in doc.paragraphs:
        text  = para.text.strip()
        style = para.style.name if para.style else ""
        if not text:
            continue

        if style.lower() in ("title", "título"):
            title = text
            continue

        if style.lower() in ("subtitle", "subtítulo", "subtitle (web)"):
            subtitle = text
            continue

        # Fallback: first non-empty line = title
        if title is None:
            title = text
            continue

        # Fallback: second non-empty short line = subtitle
        if subtitle is None and len(text) < 120:
            subtitle = text
            break

    return {
        "title":    title    or "Libro Digital",
        "subtitle": subtitle or "",
    }

# ──────────────────────────────────────────────
# 3. CONVERSIÓN DOCX → PÁGINAS
# ──────────────────────────────────────────────

def para_to_html(para) -> str:
    style = para.style.name if para.style else ""
    text  = para.text.strip()
    if not text:
        return ""
    if style.startswith("Heading 1"): return f"<h1>{html.escape(text)}</h1>"
    if style.startswith("Heading 2"): return f"<h2>{html.escape(text)}</h2>"
    if style.startswith("Heading 3"): return f"<h3>{html.escape(text)}</h3>"
    if style in ("Quote", "Intense Quote"): return f"<blockquote>{html.escape(text)}</blockquote>"
    parts = []
    for run in para.runs:
        t = html.escape(run.text)
        if run.bold:      t = f"<strong>{t}</strong>"
        if run.italic:    t = f"<em>{t}</em>"
        if run.underline: t = f"<u>{t}</u>"
        parts.append(t)
    inner = "".join(parts) or html.escape(text)
    if "List" in style:
        return f"<li>{inner}</li>"
    return f"<p>{inner}</p>"

def table_to_html(table) -> str:
    rows = []
    for i, row in enumerate(table.rows):
        cells = "".join(
            f"<{'th' if i==0 else 'td'}>{html.escape(c.text.strip())}</{'th' if i==0 else 'td'}>"
            for c in row.cells
        )
        rows.append(f"<tr>{cells}</tr>")
    return "<table>" + "".join(rows) + "</table>"

def paginate_html(content: str, max_chars: int) -> list:
    tag_pat = re.compile(
        r'(<(?:p|h[123]|blockquote|ul|li|table)[^>]*>.*?</(?:p|h[123]|blockquote|ul|table)>)',
        re.DOTALL)
    elements = [e for e in tag_pat.split(content) if e.strip()]
    chunks, current, count = [], [], 0
    for el in elements:
        length = len(re.sub(r'<[^>]+>', '', el))
        if count + length > max_chars and current:
            chunks.append("".join(current))
            current, count = [el], length
        else:
            current.append(el)
            count += length
    if current:
        chunks.append("".join(current))
    return chunks or [content]

def docx_to_pages(docx_path: str) -> tuple:
    """
    Returns (pages_list, meta_dict).
    - No generated cover — document content starts on page 1 (right side).
    - Heading 1 forces a new page, title included in content.
    - Even number of pages ensured for proper spread display.
    """
    if not HAS_DOCX:
        raise RuntimeError("Falta python-docx.")

    doc   = Document(docx_path)
    meta  = extract_meta(doc)
    pages = []
    buffer = []
    in_list = False

    # Track which paragraphs are title/subtitle (first two non-empty) to skip them
    # from the main content — they'll show in the topbar instead
    skip_indices = set()
    non_empty_count = 0
    for i, para in enumerate(doc.paragraphs):
        if para.text.strip():
            non_empty_count += 1
            if non_empty_count <= 2:
                style = para.style.name.lower() if para.style else ""
                if any(s in style for s in ("title", "subtitle", "título", "subtítulo")):
                    skip_indices.add(i)
            else:
                break

    def flush():
        if not buffer:
            return
        for chunk in paginate_html("".join(buffer), CHARS_PER_PAGE):
            pages.append({"type": "content", "content": chunk})
        buffer.clear()

    para_index = 0
    for block in doc.element.body:
        tag = block.tag.split("}")[-1] if "}" in block.tag else block.tag

        if tag == "p":
            from docx.text.paragraph import Paragraph as DocxPara
            para  = DocxPara(block, doc)
            style = para.style.name if para.style else ""
            text  = para.text.strip()

            # Skip title/subtitle paragraphs (shown in topbar)
            if para_index in skip_indices:
                para_index += 1
                continue
            para_index += 1

            if not text:
                continue

            if style.startswith("Heading 1"):
                if in_list:
                    buffer.append("</ul>")
                    in_list = False
                flush()
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
            para_index += 1

    if in_list:
        buffer.append("</ul>")
    flush()

    # Ensure even number of pages so spreads are complete
    if len(pages) % 2 != 0:
        pages.append({"type": "blank"})

    return pages, meta

# ──────────────────────────────────────────────
# 4. INYECCIÓN EN index.html
# ──────────────────────────────────────────────

def inject(pages: list, meta: dict, template_path: str, output_path: str) -> None:
    with open(template_path, "r", encoding="utf-8") as f:
        tmpl = f.read()

    out = tmpl.replace("__PAGES_JSON__", json.dumps(pages, ensure_ascii=False, indent=2))

    # Update title and subtitle in JS META object
    out = out.replace(
        'title:    "Actividades del Colegio"',
        f'title:    "{meta["title"]}"'
    )
    out = out.replace(
        'subtitle: "Registro de actividades y eventos"',
        f'subtitle: "{meta["subtitle"]}"'
    )
    # Update HTML <title>
    out = re.sub(r'<title>.*?</title>', f'<title>{meta["title"]}</title>', out)
    # Update topbar text directly too
    out = out.replace(
        '>Actividades del Colegio<',
        f'>{meta["title"]}<'
    )
    out = out.replace(
        '>Registro de actividades y eventos<',
        f'>{meta["subtitle"]}<'
    )

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(out)

    print(f"✅ Libro generado: {output_path}  ({len(pages)} páginas)")
    print(f"   Título: {meta['title']}")
    print(f"   Subtítulo: {meta['subtitle']}")

# ──────────────────────────────────────────────
# 5. MODO DEMO
# ──────────────────────────────────────────────

def demo_pages() -> tuple:
    pages = [
        {"type": "content", "content": "<h1>Día del Libro</h1><p>Los alumnos de 5° básico organizaron una feria del libro en el patio central.</p><blockquote>\"Leer es volar sin moverse del lugar.\"</blockquote>"},
        {"type": "content", "content": "<h1>Huerto Escolar</h1><p>Con apoyo de padres voluntarios, los alumnos aprendieron a sembrar tomates y lechugas.</p>"},
        {"type": "content", "content": "<h1>Olimpiadas Matemáticas</h1><table><tr><th>Lugar</th><th>Equipo</th><th>Puntos</th></tr><tr><td>1°</td><td>Los Calculadores</td><td>95</td></tr></table>"},
        {"type": "content", "content": "<h1>Festival de la Primavera</h1><p>El Festival reunió a todos los cursos para celebrar el inicio de la estación.</p>"},
    ]
    meta = {"title": "Actividades del Colegio", "subtitle": "Registro de actividades y eventos"}
    return pages, meta

# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    base_dir      = Path(__file__).parent
    template_path = base_dir / "index.html"
    output_path   = base_dir / "index.html"
    tmp_docx      = base_dir / "tmp_document.docx"

    file_id   = os.environ.get("DRIVE_FILE_ID")
    demo_mode = "--demo" in sys.argv or not file_id

    if demo_mode:
        print("ℹ️  Modo demo.")
        pages, meta = demo_pages()
    else:
        print(f"📥 Descargando documento (ID: {file_id})…")
        download_docx_from_drive(file_id, str(tmp_docx))
        print("📖 Convirtiendo…")
        pages, meta = docx_to_pages(str(tmp_docx))
        tmp_docx.unlink(missing_ok=True)

    inject(pages, meta, str(template_path), str(output_path))

if __name__ == "__main__":
    main()

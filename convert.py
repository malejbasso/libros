#!/usr/bin/env python3

import json
import os
import sys
import io
import base64
import subprocess
import tempfile
import shutil

from pathlib import Path

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

try:
    from pdf2image import convert_from_path
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

DPI = 170


# ──────────────────────────────────────────────
# DESCARGAR DOCX
# ──────────────────────────────────────────────

def download_docx(file_id: str, dest_path: str):

    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client")

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if not sa_json:
        raise RuntimeError("Falta GOOGLE_SERVICE_ACCOUNT_JSON")

    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )

    service = build(
        "drive",
        "v3",
        credentials=creds,
        cache_discovery=False
    )

    request = service.files().get_media(fileId=file_id)

    fh = io.BytesIO()

    downloader = MediaIoBaseDownload(fh, request)

    done = False

    while not done:

        status, done = downloader.next_chunk()

        if status:
            print(f"Descargando... {int(status.progress()*100)}%")

    with open(dest_path, "wb") as f:
        f.write(fh.getvalue())

    print("✅ DOCX descargado")


# ──────────────────────────────────────────────
# DOCX → PDF
# ──────────────────────────────────────────────

def docx_to_pdf(docx_path: str, output_dir: str) -> str:

    print("📄 Convirtiendo DOCX a PDF...")

    result = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            output_dir,
            docx_path
        ],
        capture_output=True,
        text=True,
        timeout=120
    )

    if result.returncode != 0:

        print(result.stdout)
        print(result.stderr)

        raise RuntimeError("LibreOffice falló")

    pdf_path = Path(output_dir) / (
        Path(docx_path).stem + ".pdf"
    )

    if not pdf_path.exists():
        raise RuntimeError("No se generó PDF")

    print("✅ PDF generado")

    return str(pdf_path)


# ──────────────────────────────────────────────
# PDF → PNG BASE64
# ──────────────────────────────────────────────

def pdf_to_images_b64(pdf_path: str):

    if not HAS_PDF2IMAGE:
        raise RuntimeError("Falta pdf2image")

    print("🖼️ Convirtiendo PDF a imágenes...")

    images = convert_from_path(
        pdf_path,
        dpi=DPI,
        fmt="png"
    )

    pages = []

    for i, img in enumerate(images):

        buf = io.BytesIO()

        img.save(
            buf,
            format="PNG",
            optimize=True
        )

        b64 = base64.b64encode(
            buf.getvalue()
        ).decode("utf-8")

        pages.append(b64)

        print(f"Página {i+1}/{len(images)}")

    return pages


# ──────────────────────────────────────────────
# GENERAR HTML
# ──────────────────────────────────────────────

def generate_html(
    pages_b64,
    title,
    subtitle,
    output_path
):

    pages_html = ""

    # Primera página en blanco
    pages_html += '<div class="page blank"></div>\n'

    # Páginas reales
    for p in pages_b64:

        pages_html += f'''
<div class="page">
    <img src="data:image/png;base64,{p}">
</div>
'''

    html = f"""
<!DOCTYPE html>
<html lang="es">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>{title}</title>

<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>

<script src="turn.js"></script>

<link rel="preconnect"
      href="https://fonts.googleapis.com">

<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&display=swap"
      rel="stylesheet">

<style>

*{{
    margin:0;
    padding:0;
    box-sizing:border-box;
}}

html,body{{
    width:100%;
    height:100%;
    overflow:hidden;

    background:#1b1008;

    font-family:'Playfair Display',serif;
}}

#toolbar{{
    position:fixed;

    top:0;
    left:0;
    right:0;

    height:60px;

    display:flex;
    align-items:center;
    justify-content:space-between;

    padding:0 30px;

    background:rgba(0,0,0,.55);

    backdrop-filter:blur(10px);

    z-index:9999;
}}

#toolbar h1{{
    color:white;
    font-size:20px;
}}

#toolbar .sub{{
    color:#c9c9c9;
    font-size:13px;
}}

#page-number{{
    color:#E0C27A;
    font-size:18px;
}}

#container{{
    width:100%;
    height:100vh;

    display:flex;
    align-items:center;
    justify-content:center;
}}

#book{{
    width:1040px;
    height:735px;
}}

.page{{
    width:520px;
    height:735px;

    background:white;

    overflow:hidden;
}}

.page img{{
    width:100%;
    height:100%;

    object-fit:contain;

    background:white;

    user-select:none;

    pointer-events:none;
}}

.blank{{
    background:white;
}}

.turn-page{{
    box-shadow:
        0 0 25px rgba(0,0,0,.25);
}}

@media(max-width:1100px){{

    #book{{
        width:340px !important;
        height:480px !important;
    }}

    .page{{
        width:340px !important;
        height:480px !important;
    }}
}}

</style>

</head>

<body>

<div id="toolbar">

    <div>
        <h1>{title}</h1>
        <div class="sub">{subtitle}</div>
    </div>

    <div id="page-number">
        Página 1
    </div>

</div>

<div id="container">

    <div id="book">

        {pages_html}

    </div>

</div>

<audio id="flipSound" preload="auto">
    <source src="page-flip.mp3" type="audio/mpeg">
</audio>

<script>

const sound =
    document.getElementById('flipSound');

const isMobile =
    window.innerWidth <= 1100;

$('#book').turn({{

    width: isMobile ? 340 : 1040,

    height: isMobile ? 480 : 735,

    autoCenter: true,

    gradients: true,

    acceleration: true,

    elevation: 50,

    duration: 900

}});

$('#book').bind('turning', function(event, page){{

    try {{
        sound.currentTime = 0;
        sound.play();
    }}
    catch(e){{}}

    let realPage = page;

    // Ajustar numeración porque existe página blanca inicial
    if(realPage > 1){{
        realPage = realPage - 1;
    }}

    document.getElementById('page-number')
        .innerText =
            'Página ' + realPage;
}});

document.addEventListener('keydown', e => {{

    if(e.key === 'ArrowRight'){{
        $('#book').turn('next');
    }}

    if(e.key === 'ArrowLeft'){{
        $('#book').turn('previous');
    }}

}});

</script>

</body>
</html>
"""

    with open(
        output_path,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(html)

    print("✅ index.html generado")


# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():

    base_dir = Path(__file__).parent

    output_path = base_dir / "index.html"

    title = os.environ.get(
        "BOOK_TITLE",
        "Libro Digital"
    )

    subtitle = os.environ.get(
        "BOOK_SUBTITLE",
        ""
    )

    file_id = os.environ.get(
        "DRIVE_FILE_ID"
    )

    if not file_id:

        print("Falta DRIVE_FILE_ID")

        sys.exit(1)

    tmp_dir = tempfile.mkdtemp()

    try:

        docx_path = os.path.join(
            tmp_dir,
            "documento.docx"
        )

        download_docx(
            file_id,
            docx_path
        )

        pdf_path = docx_to_pdf(
            docx_path,
            tmp_dir
        )

        pages_b64 = pdf_to_images_b64(
            pdf_path
        )

        generate_html(
            pages_b64,
            title,
            subtitle,
            str(output_path)
        )

    finally:

        shutil.rmtree(
            tmp_dir,
            ignore_errors=True
        )


if __name__ == "__main__":
    main()

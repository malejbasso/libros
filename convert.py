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

    pages_json = "[" + ",".join([
        '{"src":"data:image/png;base64,' + p + '"}'
        for p in pages_b64
    ]) + "]"

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

  background:#120904;

  font-family:Georgia, serif;
}}

#app{{
  width:100%;
  height:100vh;

  display:flex;
  align-items:center;
  justify-content:center;

  overflow:hidden;
}}

#flipbook{{
  box-shadow:
    0 10px 40px rgba(0,0,0,.45);
}}

.page{{
  background:white;
}}

.page img{{
  width:100%;
  height:100%;
  object-fit:contain;

  background:white;
}}

#toolbar{{
  position:fixed;

  top:0;
  left:0;
  right:0;

  height:50px;

  display:flex;
  align-items:center;
  justify-content:space-between;

  padding:0 15px;

  background:rgba(0,0,0,.7);

  color:white;

  z-index:9999;
}}

#toolbar h1{{
  font-size:16px;
}}

#toolbar small{{
  opacity:.7;
}}

#nav{{
  position:fixed;

  bottom:15px;
  left:0;
  right:0;

  display:flex;
  justify-content:center;
  align-items:center;

  gap:16px;

  z-index:9999;
}}

.btn{{
  width:48px;
  height:48px;

  border-radius:50%;
  border:none;

  background:white;

  font-size:24px;

  cursor:pointer;

  box-shadow:
    0 5px 15px rgba(0,0,0,.35);
}}

#counter{{
  color:#fff;
  font-size:18px;
}}

@media (max-width: 768px) {{

  body{{
    overflow:hidden;
  }}

  #flipbook{{
    width:88vw !important;
    height:calc(88vw * 1.414) !important;
    margin-top:10px;
  }}

  .page img{{
    width:100%;
    height:100%;
    object-fit:contain;
  }}

  #toolbar{{
    height:44px;
    padding:0 10px;
  }}

  #toolbar h1{{
    font-size:13px;
  }}

  #pageLabel{{
    font-size:12px;
  }}

  #nav{{
    bottom:12px;
    gap:12px;
  }}

  .btn{{
    width:42px;
    height:42px;
    font-size:20px;
  }}

  #counter{{
    font-size:16px;
  }}
}}

</style>

</head>

<body>

<div id="toolbar">

  <div>
    <h1>{title}</h1>
    <small>{subtitle}</small>
  </div>

  <div id="pageLabel">
    Página 1
  </div>

</div>

<div id="app">

  <div id="flipbook"></div>

</div>

<div id="nav">

  <button class="btn" id="prev">
    ←
  </button>

  <div id="counter"></div>

  <button class="btn" id="next">
    →
  </button>

</div>

<audio id="pageSound" preload="auto">
  <source src="page-flip.mp3" type="audio/mpeg">
</audio>

<script>

const rawPages = {pages_json};

const isMobile = window.innerWidth <= 768;

// Desktop:
// [vacío][1]
// [2][3]
// [4][5]

const pages = isMobile
  ? rawPages
  : [null, ...rawPages];

const flipbook =
  document.getElementById('flipbook');

for(let i = 0; i < pages.length; i++){{

  const pageData = pages[i];

  const div = document.createElement('div');

  div.className = 'page';

  if(pageData){{

    const img = document.createElement('img');

    img.src = pageData.src;

    div.appendChild(img);

  }}else{{

    div.style.background = 'white';
  }}

  flipbook.appendChild(div);
}}

$('#flipbook').turn({{

  width: isMobile
    ? window.innerWidth * 0.88
    : 1100,

  height: isMobile
    ? (window.innerWidth * 0.88) * 1.414
    : 780,

  autoCenter: true,

  display: isMobile
    ? 'single'
    : 'double',

  elevation: 50,

  gradients: true,

  acceleration: true,

  when: {{

    turning: function(event, page){{

      const audio =
        document.getElementById('pageSound');

      audio.currentTime = 0;

      audio.play().catch(()=>{{}});
    }},

    turned: function(event, page){{

      document.getElementById('pageLabel')
        .textContent = 'Página ' + page;

      document.getElementById('counter')
        .textContent =
          page + ' / ' + pages.length;
    }}
  }}
}});

if(isMobile){{
  $('#flipbook').turn('page', 2);
}}

document.getElementById('prev')
  .addEventListener('click', ()=>{{
    $('#flipbook').turn('previous');
  }});

document.getElementById('next')
  .addEventListener('click', ()=>{{
    $('#flipbook').turn('next');
  }});

document.addEventListener('keydown', e=>{{

  if(e.key === 'ArrowRight'){{
    $('#flipbook').turn('next');
  }}

  if(e.key === 'ArrowLeft'){{
    $('#flipbook').turn('previous');
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

```python
#!/usr/bin/env python3
"""
convert.py — Descarga el .docx desde Google Drive, lo convierte a PDF
             con LibreOffice, y genera el libro digital con las páginas
             como imágenes.

Dependencias Python:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pdf2image pillow

Dependencias sistema:
    sudo apt-get install -y libreoffice poppler-utils
"""

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

DPI = 150


# ──────────────────────────────────────────────
# DESCARGAR DOCX DESDE GOOGLE DRIVE
# ──────────────────────────────────────────────

def download_docx(file_id: str, dest_path: str) -> None:

    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client")

    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")

    if not sa_json:
        raise ValueError("Variable GOOGLE_SERVICE_ACCOUNT_JSON no encontrada")

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

    meta = service.files().get(
        fileId=file_id,
        fields="name,mimeType"
    ).execute()

    print(f"📄 Archivo: {meta.get('name')}")

    request = service.files().get_media(fileId=file_id)

    fh = io.BytesIO()

    downloader = MediaIoBaseDownload(fh, request)

    done = False

    while not done:

        status, done = downloader.next_chunk()

        if status:
            print(f"Descargando... {int(status.progress() * 100)}%")

    with open(dest_path, "wb") as f:
        f.write(fh.getvalue())

    print(f"✅ Descargado: {dest_path}")


# ──────────────────────────────────────────────
# DOCX → PDF
# ──────────────────────────────────────────────

def docx_to_pdf(docx_path: str, output_dir: str) -> str:

    print("📄 Convirtiendo a PDF...")

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

    docx_name = Path(docx_path).stem

    pdf_path = Path(output_dir) / f"{docx_name}.pdf"

    if not pdf_path.exists():

        pdfs = list(Path(output_dir).glob("*.pdf"))

        if not pdfs:
            raise RuntimeError("No se generó PDF")

        pdf_path = pdfs[0]

    print(f"✅ PDF generado: {pdf_path}")

    return str(pdf_path)


# ──────────────────────────────────────────────
# PDF → IMÁGENES BASE64
# ──────────────────────────────────────────────

def pdf_to_images_b64(pdf_path: str, dpi: int = DPI) -> list:

    if not HAS_PDF2IMAGE:
        raise RuntimeError("Falta pdf2image")

    print("🖼️ Convirtiendo PDF a imágenes...")

    images = convert_from_path(
        pdf_path,
        dpi=dpi,
        fmt="png"
    )

    print(f"{len(images)} páginas encontradas")

    b64_pages = []

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

        b64_pages.append(b64)

        print(f"Página {i+1}/{len(images)}")

    return b64_pages


# ──────────────────────────────────────────────
# GENERAR HTML
# ──────────────────────────────────────────────

def generate_html(
    pages_b64: list,
    title: str,
    subtitle: str,
    output_path: str
):

    if not pages_b64:
        raise RuntimeError("No hay páginas")

    parts = []

    for b64 in pages_b64:

        parts.append(
            '{"type":"img","src":"data:image/png;base64,' + b64 + '"}'
        )

    pages_json = "[" + ",".join(parts) + "]"

    html = f"""<!DOCTYPE html>
<html lang="es">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>{title}</title>

<link rel="preconnect"
      href="https://fonts.googleapis.com">

<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Lora:wght@400;500&display=swap"
      rel="stylesheet">

<style>

*,*::before,*::after{{
  box-sizing:border-box;
  margin:0;
  padding:0;
}}

:root{{
  --gold:#B8862A;
  --gold2:#D4A843;
  --gold3:#F0C96A;
  --cream:#F7F0E0;

  --pw:480px;
  --ph:660px;
}}

html,body{{
  height:100%;
  overflow:hidden;
  background:#1A0F08;
  font-family:'Lora',serif;
}}

#stage{{
  height:100vh;

  display:flex;
  align-items:center;
  justify-content:center;

  perspective:2200px;
}}

#book-wrap{{
  position:relative;

  width:calc(var(--pw) * 2);
  height:var(--ph);

  transform-style:preserve-3d;
}}

.page{{
  position:absolute;

  width:var(--pw);
  height:var(--ph);

  top:0;

  background:var(--cream);

  overflow:hidden;

  backface-visibility:hidden;
}}

.left-page{{
  left:0;
}}

.right-page{{
  left:var(--pw);
}}

.page img{{
  width:100%;
  height:100%;
  object-fit:contain;
}}

.blank-page{{
  background:var(--cream);
}}

#flip-layer{{
  position:absolute;

  width:var(--pw);
  height:var(--ph);

  top:0;

  transform-style:preserve-3d;

  pointer-events:none;

  z-index:100;
}}

#flip-front,
#flip-back{{
  position:absolute;
  inset:0;

  backface-visibility:hidden;
}}

#flip-back{{
  transform:rotateY(180deg);
}}

#topbar{{
  position:fixed;

  top:0;
  left:0;
  right:0;

  height:52px;

  z-index:1000;

  display:flex;
  align-items:center;
  justify-content:space-between;

  padding:0 24px;

  background:rgba(0,0,0,.85);

  color:white;
}}

#nav{{
  position:fixed;

  bottom:20px;
  left:0;
  right:0;

  display:flex;
  align-items:center;
  justify-content:center;

  gap:20px;
}}

.nav-btn{{
  width:46px;
  height:46px;

  border:none;
  border-radius:50%;

  cursor:pointer;

  font-size:22px;
}}

@media(max-width:980px){{

  :root{{
    --pw:300px;
    --ph:420px;
  }}

  #book-wrap{{
    width:var(--pw);
  }}

  .left-page{{
    display:none;
  }}
}}

</style>
</head>

<body>

<div id="topbar">

  <div>
    <h1>{title}</h1>
    <div>{subtitle}</div>
  </div>

  <div id="page-counter">
    Pág. 1
  </div>

</div>

<div id="stage">

  <div id="book-wrap">

    <div id="right-slot"></div>

    <div id="left-slot"></div>

    <div id="flip-layer">

      <div id="flip-front"></div>

      <div id="flip-back"></div>

    </div>

  </div>

</div>

<div id="nav">

  <button class="nav-btn" id="btn-prev">
    ←
  </button>

  <span id="page-info"></span>

  <button class="nav-btn" id="btn-next">
    →
  </button>

</div>

<script>

const PAGES={pages_json};

function mkPage(data,side){{

  const d=document.createElement('div');

  d.className='page '+side;

  if(!data){{

    d.classList.add('blank-page');

    return d;
  }}

  const img=document.createElement('img');

  img.src=data.src;

  d.appendChild(img);

  return d;
}}

const isMobile=window.innerWidth<=980;

const totalSpreads=Math.ceil((PAGES.length+1)/2);

let cur=0;
let busy=false;

const RS=document.getElementById('right-slot');
const LS=document.getElementById('left-slot');

const FL=document.getElementById('flip-layer');
const FF=document.getElementById('flip-front');
const FB=document.getElementById('flip-back');

function sp(n){{

  if(n===0){{
    return {{
      l:null,
      r:PAGES[0]||null
    }};
  }}

  const leftIndex=(n*2)-1;
  const rightIndex=(n*2);

  return {{
    l:PAGES[leftIndex]||null,
    r:PAGES[rightIndex]||null
  }};
}}

function render(n){{

  const s=sp(n);

  RS.innerHTML='';
  LS.innerHTML='';

  if(!isMobile){{
    LS.appendChild(
      mkPage(s.l,'left-page')
    );
  }}

  RS.appendChild(
    mkPage(s.r,'right-page')
  );
}}

function ui(){{

  let currentPage;

  if(cur===0){{
    currentPage=1;
  }}else{{
    currentPage=cur*2;
  }}

  document.getElementById('page-counter')
    .textContent=`Pág. ${{currentPage}}`;

  document.getElementById('page-info')
    .textContent=`${{cur+1}} / ${{totalSpreads}}`;

  document.getElementById('btn-prev')
    .disabled=cur<=0;

  document.getElementById('btn-next')
    .disabled=cur>=totalSpreads-1;
}}

function go(dir){{

  if(busy)return;

  const nx=cur+dir;

  if(nx<0 || nx>=totalSpreads)return;

  busy=true;

  const cs=sp(cur);
  const ns=sp(nx);

  if(isMobile){{

    cur=nx;

    render(nx);

    ui();

    busy=false;

    return;
  }}

  if(dir>0){{

    FL.style.left='var(--pw)';
    FL.style.transformOrigin='left center';

    FF.innerHTML='';
    FB.innerHTML='';

    if(cs.r){{

      const e=mkPage(cs.r,'right-page');

      e.style.cssText='position:absolute;inset:0';

      FF.appendChild(e);
    }}

    if(ns.l){{

      const e=mkPage(ns.l,'left-page');

      e.style.cssText='position:absolute;inset:0';

      FB.appendChild(e);
    }}

  }}else{{

    FL.style.left='0';
    FL.style.transformOrigin='right center';

    FF.innerHTML='';
    FB.innerHTML='';

    if(cs.l){{

      const e=mkPage(cs.l,'left-page');

      e.style.cssText='position:absolute;inset:0';

      FF.appendChild(e);
    }}

    if(ns.r){{

      const e=mkPage(ns.r,'right-page');

      e.style.cssText='position:absolute;inset:0';

      FB.appendChild(e);
    }}
  }}

  FL.style.transition='none';

  FL.style.transform='rotateY(0deg)';

  const deg=dir>0?-180:180;

  requestAnimationFrame(()=>{{
    requestAnimationFrame(()=>{{

      FL.style.transition=
        'transform 480ms cubic-bezier(.645,.045,.355,1)';

      FL.style.transform=
        `rotateY(${{deg}}deg)`;

    }});
  }});

  setTimeout(()=>{{

    cur=nx;

    render(nx);

    ui();

    FL.style.transition='none';

    FL.style.transform='rotateY(0deg)';

    busy=false;

  }},520);
}}

document.getElementById('btn-prev')
  .addEventListener('click',()=>go(-1));

document.getElementById('btn-next')
  .addEventListener('click',()=>go(1));

document.addEventListener('keydown',e=>{{

  if(e.key==='ArrowRight')go(1);

  if(e.key==='ArrowLeft')go(-1);
}});

render(0);

ui();

</script>

</body>
</html>
"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"✅ HTML generado: {output_path}")


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

    file_id = os.environ.get("DRIVE_FILE_ID")

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
            pdf_path,
            dpi=DPI
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
```

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

<link rel="preconnect"
      href="https://fonts.googleapis.com">

<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=Lora:wght@400;500&display=swap"
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

  background:#120904;

  font-family:'Lora',serif;
}}

:root{{
  --pw:520px;
  --ph:735px;
}}

#stage{{
  width:100%;
  height:100vh;

  display:flex;
  align-items:center;
  justify-content:center;

  perspective:2400px;
}}

#book{{
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

  background:white;

  overflow:hidden;

  box-shadow:
    0 0 30px rgba(0,0,0,.15);

  border-radius:2px;
}}

.left{{
  left:0;
}}

.right{{
  left:var(--pw);
}}

.page img{{
  width:100%;
  height:100%;
  object-fit:contain;

  background:white;
}}

#book::after{{
  content:'';

  position:absolute;

  left:50%;
  top:0;

  width:30px;
  height:100%;

  transform:translateX(-50%);

  background:
    linear-gradient(
      to right,
      rgba(0,0,0,.18),
      rgba(0,0,0,.05),
      rgba(255,255,255,.6),
      rgba(0,0,0,.05),
      rgba(0,0,0,.18)
    );

  filter:blur(2px);

  z-index:30;

  pointer-events:none;
}}

#flip{{
  position:absolute;

  width:var(--pw);
  height:var(--ph);

  top:0;

  transform-style:preserve-3d;

  z-index:100;

  pointer-events:none;
}}

#front,
#back{{
  position:absolute;
  inset:0;

  backface-visibility:hidden;
}}

#back{{
  transform:rotateY(180deg);
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

  background:rgba(0,0,0,.65);

  backdrop-filter:blur(10px);

  z-index:500;
}}

#toolbar h1{{
  color:white;
  font-size:18px;
}}

#toolbar div{{
  color:#ddd;
}}

#nav{{
  position:fixed;

  bottom:25px;
  left:0;
  right:0;

  display:flex;
  align-items:center;
  justify-content:center;

  gap:22px;

  z-index:999;
}}

.btn{{
  width:56px;
  height:56px;

  border-radius:50%;
  border:none;

  background:white;

  font-size:24px;

  cursor:pointer;

  box-shadow:
    0 5px 20px rgba(0,0,0,.35);
}}

#counter{{
  color:#D9B45A;
  font-size:22px;
}}

@media(max-width:1100px){{

  :root{{
    --pw:340px;
    --ph:480px;
  }}

  #book{{
    width:var(--pw);
  }}

  .left{{
    display:none;
  }}

  #book::after{{
    display:none;
  }}
}}

</style>
</head>

<body>

<div id="toolbar">

  <div>
    <h1>{title}</h1>
    <div>{subtitle}</div>
  </div>

  <div id="pageLabel">
    Página 1
  </div>

</div>

<div id="stage">

  <div id="book">

    <div id="leftSlot"></div>

    <div id="rightSlot"></div>

    <div id="flip">

      <div id="front"></div>

      <div id="back"></div>

    </div>

  </div>

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

<script>

const PAGES = {pages_json};

const isMobile =
  window.innerWidth <= 1100;


// ─────────────────────────────
// ORDEN REAL DE LIBRO
// ─────────────────────────────
//
// Spread 0 = [vacío][1]
//
// Spread 1 = [2][3]
//
// Spread 2 = [4][5]
//
// Spread 3 = [6][7]
//
// ─────────────────────────────

function spread(n){{

  // Primera vista
  if(n === 0){{
    return {{
      left:null,
      right:PAGES[0]
    }};
  }}

  // Spread reales
  const leftIndex =
    (n * 2);

  const rightIndex =
    (n * 2) + 1;

  return {{
    left:PAGES[leftIndex] || null,
    right:PAGES[rightIndex] || null
  }};
}}

const total =
  Math.ceil((PAGES.length + 1)/2);

let current = 0;

let busy = false;

const leftSlot =
  document.getElementById('leftSlot');

const rightSlot =
  document.getElementById('rightSlot');

const flip =
  document.getElementById('flip');

const front =
  document.getElementById('front');

const back =
  document.getElementById('back');

function makePage(data, cls){{

  const d = document.createElement('div');

  d.className = 'page ' + cls;

  if(!data){{
    return d;
  }}

  const img = document.createElement('img');

  img.src = data.src;

  d.appendChild(img);

  return d;
}}

function render(n){{

  const s = spread(n);

  leftSlot.innerHTML = '';
  rightSlot.innerHTML = '';

  if(!isMobile){{
    leftSlot.appendChild(
      makePage(s.left,'left')
    );
  }}

  rightSlot.appendChild(
    makePage(s.right,'right')
  );
}}

function updateUI(){{

  let p;

  if(current === 0){{
    p = 1;
  }} else {{
    p = current * 2;
  }}

  document.getElementById('pageLabel')
    .textContent = 'Página ' + p;

  document.getElementById('counter')
    .textContent =
      (current+1) + ' / ' + total;
}}

function go(dir){{

  if(busy)return;

  const next = current + dir;

  if(next < 0 || next >= total)return;

  busy = true;

  const cs = spread(current);
  const ns = spread(next);

  if(isMobile){{

    current = next;

    render(current);

    updateUI();

    busy = false;

    return;
  }}

  if(dir > 0){{

    flip.style.left='var(--pw)';
    flip.style.transformOrigin='left center';

    front.innerHTML='';
    back.innerHTML='';

    if(cs.right){{

      const p = makePage(cs.right,'right');

      p.style.cssText =
        'position:absolute;inset:0';

      front.appendChild(p);
    }}

    if(ns.left){{

      const p = makePage(ns.left,'left');

      p.style.cssText =
        'position:absolute;inset:0';

      back.appendChild(p);
    }}

  }} else {{

    flip.style.left='0';
    flip.style.transformOrigin='right center';

    front.innerHTML='';
    back.innerHTML='';

    if(cs.left){{

      const p = makePage(cs.left,'left');

      p.style.cssText =
        'position:absolute;inset:0';

      front.appendChild(p);
    }}

    if(ns.right){{

      const p = makePage(ns.right,'right');

      p.style.cssText =
        'position:absolute;inset:0';

      back.appendChild(p);
    }}
  }}

  leftSlot.innerHTML='';
  rightSlot.innerHTML='';

  leftSlot.appendChild(
    makePage(ns.left,'left')
  );

  rightSlot.appendChild(
    makePage(ns.right,'right')
  );

  flip.style.transition='none';

  flip.style.transform='rotateY(0deg)';

  requestAnimationFrame(()=>{{
    requestAnimationFrame(()=>{{

      flip.style.transition =
        'transform 550ms cubic-bezier(.645,.045,.355,1)';

      flip.style.transform =
        dir > 0
          ? 'rotateY(-180deg)'
          : 'rotateY(180deg)';

    }});
  }});

  setTimeout(()=>{{

    flip.style.transition='none';

    flip.style.transform='rotateY(0deg)';

    current = next;

    render(current);

    updateUI();

    busy = false;

  }},580);
}}

document.getElementById('prev')
  .addEventListener('click',()=>go(-1));

document.getElementById('next')
  .addEventListener('click',()=>go(1));

document.addEventListener('keydown',e=>{{

  if(e.key==='ArrowRight')go(1);

  if(e.key==='ArrowLeft')go(-1);

}});

render(0);

updateUI();

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

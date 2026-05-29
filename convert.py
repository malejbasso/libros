#!/usr/bin/env python3
"""
convert.py — Descarga el .docx desde Google Drive como PDF,
             convierte cada página a imagen y genera el libro digital.

Dependencias:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pdf2image pillow

También requiere poppler instalado en el sistema:
    Ubuntu/Debian: sudo apt-get install -y poppler-utils
    Mac:           brew install poppler
    (En GitHub Actions se instala automáticamente, ver publish.yml)

Variables de entorno:
    GOOGLE_SERVICE_ACCOUNT_JSON  — JSON de la cuenta de servicio
    DRIVE_FILE_ID                — ID del .docx en Google Drive
    BOOK_TITLE                   — Título (opcional, se toma del doc si no se define)
    BOOK_SUBTITLE                — Subtítulo (opcional)
"""

import json, os, sys, io, base64, re
from pathlib import Path

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    HAS_GOOGLE = True
except ImportError:
    HAS_GOOGLE = False

try:
    from pdf2image import convert_from_bytes
    HAS_PDF2IMAGE = True
except ImportError:
    HAS_PDF2IMAGE = False

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
DPI = 150  # calidad de imagen (150 = buena calidad, tamaño razonable)
           # subir a 200 para más nitidez, bajar a 120 para archivos más livianos

# ──────────────────────────────────────────────
# 1. DESCARGA DESDE GOOGLE DRIVE COMO PDF
# ──────────────────────────────────────────────

def get_drive_service():
    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client.")
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise ValueError("Variable GOOGLE_SERVICE_ACCOUNT_JSON no encontrada.")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)

def download_as_pdf(file_id: str) -> bytes:
    """
    Exporta el archivo de Drive como PDF y retorna los bytes.
    Funciona con .docx, .doc, Google Docs, .pptx, etc.
    """
    service = get_drive_service()

    # Primero detectar el tipo de archivo
    meta = service.files().get(fileId=file_id, fields="mimeType,name").execute()
    mime = meta.get("mimeType", "")
    name = meta.get("name", "")
    print(f"📄 Archivo: {name} ({mime})")

    if mime == "application/pdf":
        # Ya es PDF, descargar directo
        request = service.files().get_media(fileId=file_id)
    else:
        # Exportar como PDF (funciona para Docs, Sheets, Slides, docx, etc.)
        request = service.files().export_media(
            fileId=file_id,
            mimeType="application/pdf"
        )

    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"   Descargando... {int(status.progress() * 100)}%")

    print(f"✅ PDF descargado ({len(fh.getvalue()) // 1024} KB)")
    return fh.getvalue()

# ──────────────────────────────────────────────
# 2. CONVERTIR PDF → IMÁGENES BASE64
# ──────────────────────────────────────────────

def pdf_to_images_b64(pdf_bytes: bytes, dpi: int = DPI) -> list:
    """
    Convierte cada página del PDF a PNG base64.
    Retorna lista de strings base64.
    """
    if not HAS_PDF2IMAGE:
        raise RuntimeError("Falta pdf2image. Ejecuta: pip install pdf2image pillow")

    print(f"🖼️  Convirtiendo páginas a imágenes (DPI={dpi})...")
    images = convert_from_bytes(pdf_bytes, dpi=dpi, fmt="png")
    print(f"   {len(images)} páginas encontradas")

    b64_pages = []
    for i, img in enumerate(images):
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        b64_pages.append(b64)
        print(f"   Página {i+1}/{len(images)} convertida ({len(b64)//1024} KB)")

    return b64_pages

# ──────────────────────────────────────────────
# 3. GENERAR index.html
# ──────────────────────────────────────────────

def generate_html(pages_b64: list, title: str, subtitle: str, output_path: str):
    """
    Genera el index.html completo con las imágenes embebidas.
    No depende de ningún template externo.
    """

    # Build JS pages array
    pages_js = []
    # First spread: page 1 alone on the right (left side blank)
    pages_js.append('{"type":"blank"}')          # PAGES[0] right -> blank (placeholder)
    # Actually: PAGES[0]=page1(right), PAGES[1]=blank(left)
    # So: insert real pages first, then blank at index 1

    real_pages = [f'{{"type":"img","src":"data:image/png;base64,{b64}"}}' for b64 in pages_b64]

    # Build final array: [page1, blank, page2, page3, page4, ...]
    # page1 alone on right of spread 0, blank on left
    if real_pages:
        final = [real_pages[0], '{"type":"blank"}'] + real_pages[1:]
    else:
        final = []

    # Ensure even count
    if len(final) % 2 != 0:
        final.append('{"type":"blank"}')

    pages_json = "[\n  " + ",\n  ".join(final) + "\n]"

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{title}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,700;1,400&family=Lora:ital,wght@0,400;0,500;1,400&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --brown:#3D2B1F;--gold:#B8862A;--gold2:#D4A843;--gold3:#F0C96A;
  --muted:#8A6F52;--cream:#F7F0E0;
  --pw:480px;--ph:660px;
}}
html,body{{height:100%;background:#1A0F08;font-family:'Lora',Georgia,serif;overflow:hidden;}}
body::before{{
  content:'';position:fixed;inset:0;z-index:0;
  background:
    repeating-linear-gradient(92deg,transparent,transparent 2px,rgba(255,255,255,.012) 2px,rgba(255,255,255,.012) 4px),
    radial-gradient(ellipse at 30% 40%,#3D2010 0%,#1A0F08 60%),
    radial-gradient(ellipse at 70% 70%,#2A1608 0%,#120A04 70%);
  pointer-events:none;
}}
#topbar{{
  position:fixed;top:0;left:0;right:0;z-index:200;height:52px;
  background:rgba(26,15,8,.96);border-bottom:1px solid rgba(184,134,42,.25);
  display:flex;align-items:center;justify-content:space-between;padding:0 28px;
}}
#topbar h1{{font-family:'Playfair Display',serif;font-size:16px;font-weight:700;color:var(--gold3);letter-spacing:.06em;}}
#topbar .sub{{font-size:11px;color:var(--muted);font-style:italic;margin-top:1px}}
#page-counter{{font-size:12px;color:var(--gold2);font-family:'Playfair Display',serif;letter-spacing:.08em;}}
#stage{{
  position:relative;z-index:1;height:100vh;
  display:flex;align-items:center;justify-content:center;
  padding-top:52px;perspective:2200px;perspective-origin:50% 45%;
}}
#book-wrap{{
  position:relative;
  width:calc(var(--pw)*2);height:var(--ph);
  transform-style:preserve-3d;
  transform:rotateX(2deg);
  filter:drop-shadow(0 32px 64px rgba(0,0,0,.8)) drop-shadow(0 8px 16px rgba(0,0,0,.6));
}}
#book-wrap::after{{
  content:'';position:absolute;left:50%;top:0;bottom:0;width:4px;margin-left:-2px;
  background:linear-gradient(to right,rgba(26,15,8,.7),rgba(90,55,30,.5),rgba(26,15,8,.7));
  z-index:50;pointer-events:none;
}}
.page{{
  position:absolute;width:var(--pw);height:var(--ph);top:0;
  background:var(--cream);backface-visibility:hidden;overflow:hidden;
}}
.page.left-page{{left:0;transform-origin:right center;}}
.page.right-page{{left:var(--pw);transform-origin:left center;}}
.page.left-page::after{{
  content:'';position:absolute;top:0;right:0;width:24px;height:100%;
  background:linear-gradient(to left,rgba(30,18,8,.12),transparent);z-index:2;pointer-events:none;
}}
.page.right-page::before{{
  content:'';position:absolute;top:0;left:0;width:24px;height:100%;
  background:linear-gradient(to right,rgba(30,18,8,.08),transparent);z-index:2;pointer-events:none;
}}
.page img{{
  width:100%;height:100%;
  object-fit:contain;
  object-position:center top;
  display:block;
}}
.page.blank-page{{background:var(--cream);}}
#flip-layer{{
  position:absolute;width:var(--pw);height:var(--ph);top:0;
  transform-style:preserve-3d;pointer-events:none;z-index:100;
}}
#flip-front,#flip-back{{position:absolute;inset:0;backface-visibility:hidden;overflow:hidden;}}
#flip-back{{transform:rotateY(180deg);}}
#flip-shadow{{
  position:absolute;inset:0;z-index:10;pointer-events:none;opacity:0;
}}
#nav{{
  position:fixed;bottom:24px;left:0;right:0;z-index:200;
  display:flex;align-items:center;justify-content:center;gap:20px;
}}
.nav-btn{{
  background:rgba(26,15,8,.9);border:1px solid rgba(184,134,42,.35);
  color:var(--gold2);width:46px;height:46px;border-radius:50%;
  font-size:20px;cursor:pointer;display:flex;align-items:center;justify-content:center;
  transition:all .2s;position:relative;
}}
.nav-btn:hover:not(:disabled){{background:rgba(61,43,31,.95);border-color:var(--gold);color:var(--gold3);}}
.nav-btn:active:not(:disabled){{transform:scale(.93);}}
.nav-btn:disabled{{opacity:.25;cursor:not-allowed;}}
.nav-btn .hint{{
  position:absolute;bottom:-20px;font-size:10px;color:var(--muted);
  font-family:'Lora',serif;font-style:italic;white-space:nowrap;
}}
#page-info{{font-size:13px;color:var(--gold2);font-family:'Playfair Display',serif;min-width:80px;text-align:center;}}
#loading{{
  position:fixed;inset:0;z-index:300;background:#1A0F08;
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;
  transition:opacity .6s;
}}
#loading.hidden{{opacity:0;pointer-events:none;}}
.spinner{{
  width:36px;height:36px;border:2px solid rgba(184,134,42,.2);border-top-color:var(--gold);
  border-radius:50%;animation:spin .85s linear infinite;
}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
#loading p{{font-family:'Playfair Display',serif;color:var(--gold2);font-style:italic;font-size:15px;}}
@media(max-width:960px){{
  :root{{--pw:320px;--ph:450px;}}
  #book-wrap{{width:var(--pw)!important;}}
  .page.left-page{{display:none;}}
  #book-wrap::after{{display:none;}}
}}
</style>
</head>
<body>
<div id="loading"><div class="spinner"></div><p>Preparando el libro...</p></div>
<div id="topbar">
  <div>
    <h1>{title}</h1>
    <div class="sub">{subtitle}</div>
  </div>
  <div id="page-counter">Pág. 1</div>
</div>
<div id="stage">
  <div id="book-wrap">
    <div id="right-slot"></div>
    <div id="left-slot"></div>
    <div id="flip-layer">
      <div id="flip-front"><div id="flip-shadow"></div></div>
      <div id="flip-back"></div>
    </div>
  </div>
</div>
<div id="nav">
  <button class="nav-btn" id="btn-prev">&#8592;<span class="hint">anterior</span></button>
  <span id="page-info">1 / 1</span>
  <button class="nav-btn" id="btn-next">&#8594;<span class="hint">siguiente</span></button>
</div>
<script>
const PAGES = {pages_json};

let audioCtx = null;
function getAudio() {{
  if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
  return audioCtx;
}}
function playPageFlip() {{
  try {{
    const ctx = getAudio();
    const buf = ctx.createBuffer(1, ctx.sampleRate * 0.22, ctx.sampleRate);
    const data = buf.getChannelData(0);
    const sr = ctx.sampleRate;
    for (let i = 0; i < data.length; i++) {{
      const t = i / sr;
      const s = (Math.random() * 2 - 1);
      const env = Math.exp(-t * 18) * (1 - Math.exp(-t * 180));
      const thump = Math.exp(-t * 40) * Math.sin(2 * Math.PI * 80 * t) * 0.35;
      const crinkle = s * Math.sin(2 * Math.PI * (800 + 400 * Math.sin(t * 60)) * t) * 0.5;
      data[i] = (s * env * 0.4 + crinkle * env * 0.4 + thump) * 0.55;
    }}
    const src = ctx.createBufferSource(); src.buffer = buf;
    const bpf = ctx.createBiquadFilter(); bpf.type='bandpass'; bpf.frequency.value=1800; bpf.Q.value=0.6;
    const hpf = ctx.createBiquadFilter(); hpf.type='highpass'; hpf.frequency.value=300;
    const gain = ctx.createGain(); gain.gain.value=0.55;
    src.connect(hpf); hpf.connect(bpf); bpf.connect(gain); gain.connect(ctx.destination);
    src.start();
  }} catch(e) {{}}
}}

function buildPageEl(pageData, side) {{
  const div = document.createElement('div');
  div.className = 'page ' + side;
  if (!pageData || pageData.type === 'blank') {{
    div.classList.add('blank-page');
    return div;
  }}
  if (pageData.type === 'img') {{
    const img = document.createElement('img');
    img.src = pageData.src;
    img.alt = '';
    img.draggable = false;
    div.appendChild(img);
  }}
  return div;
}}

// Spread N: PAGES[N*2]=right(impar), PAGES[N*2+1]=left(par)
const isMobile = window.innerWidth <= 960;
const totalSpreads = Math.ceil(PAGES.length / 2);
let currentSpread = 0, animating = false;

const leftSlot  = document.getElementById('left-slot');
const rightSlot = document.getElementById('right-slot');
const flipLayer = document.getElementById('flip-layer');
const flipFront = document.getElementById('flip-front');
const flipBack  = document.getElementById('flip-back');
const flipShadow= document.getElementById('flip-shadow');

function getSpread(n) {{
  return {{ rPage: PAGES[n*2]||null, lPage: PAGES[n*2+1]||null, ri:n*2, li:n*2+1 }};
}}

function renderSpread(n) {{
  const {{rPage,lPage}} = getSpread(n);
  rightSlot.innerHTML = '';
  leftSlot.innerHTML  = '';
  rightSlot.appendChild(buildPageEl(rPage, 'right-page'));
  if (!isMobile) leftSlot.appendChild(buildPageEl(lPage, 'left-page'));
}}

const FLIP_MS = 480;
function flipTo(dir) {{
  if (animating) return;
  const next = currentSpread + dir;
  if (next < 0 || next >= totalSpreads) return;
  animating = true;
  playPageFlip();
  const cur = getSpread(currentSpread);
  const nxt = getSpread(next);

  if (isMobile) {{
    const f = rightSlot.firstChild;
    if (f) {{ f.style.transition='opacity .25s'; f.style.opacity='0'; }}
    setTimeout(() => {{
      currentSpread=next; renderSpread(next);
      const t=rightSlot.firstChild;
      if(t){{t.style.opacity='0';t.style.transition='opacity .25s';}}
      requestAnimationFrame(()=>{{ if(t) t.style.opacity='1'; }});
      updateUI(); animating=false;
    }}, 250);
    return;
  }}

  if (dir > 0) {{
    // Página derecha actual gira a la izquierda
    flipLayer.style.left='var(--pw)'; flipLayer.style.transformOrigin='left center';
    flipFront.innerHTML='';
    flipShadow.style.background='linear-gradient(to left,rgba(0,0,0,.35),transparent 70%)';
    flipFront.appendChild(flipShadow);
    if (cur.rPage) {{ const e=buildPageEl(cur.rPage,'right-page'); e.style.cssText='position:absolute;inset:0'; flipFront.appendChild(e); }}
    flipBack.innerHTML='';
    if (nxt.lPage) {{ const e=buildPageEl(nxt.lPage,'left-page'); e.style.cssText='position:absolute;inset:0;transform:scaleX(-1)'; flipBack.appendChild(e); }}
    rightSlot.innerHTML='';
    rightSlot.appendChild(buildPageEl(nxt.rPage,'right-page'));
  }} else {{
    // Página izquierda actual gira a la derecha
    flipLayer.style.left='0'; flipLayer.style.transformOrigin='right center';
    flipFront.innerHTML='';
    flipShadow.style.background='linear-gradient(to right,rgba(0,0,0,.35),transparent 70%)';
    flipFront.appendChild(flipShadow);
    if (cur.lPage) {{ const e=buildPageEl(cur.lPage,'left-page'); e.style.cssText='position:absolute;inset:0'; flipFront.appendChild(e); }}
    flipBack.innerHTML='';
    if (nxt.rPage) {{ const e=buildPageEl(nxt.rPage,'right-page'); e.style.cssText='position:absolute;inset:0;transform:scaleX(-1)'; flipBack.appendChild(e); }}
    leftSlot.innerHTML='';
    leftSlot.appendChild(buildPageEl(nxt.lPage,'left-page'));
  }}

  flipLayer.style.transition='none'; flipLayer.style.transform='rotateY(0deg)'; flipShadow.style.opacity='0';
  const deg = dir>0 ? -180 : 180;
  requestAnimationFrame(()=>requestAnimationFrame(()=>{{
    flipLayer.style.transition=`transform ${{FLIP_MS}}ms cubic-bezier(.645,.045,.355,1)`;
    flipShadow.style.transition=`opacity ${{FLIP_MS}}ms`;
    flipLayer.style.transform=`rotateY(${{deg}}deg)`;
    flipShadow.style.opacity='1';
  }}));
  setTimeout(()=>{{
    flipLayer.style.transition='none'; flipLayer.style.transform='rotateY(0deg)'; flipShadow.style.opacity='0';
    currentSpread=next; renderSpread(next); updateUI(); animating=false;
  }}, FLIP_MS+30);
}}

function updateUI() {{
  const p = currentSpread*2+1;
  document.getElementById('page-counter').textContent=`Pág. ${{p}}`;
  document.getElementById('page-info').textContent=`${{currentSpread+1}} / ${{totalSpreads}}`;
  document.getElementById('btn-prev').disabled = currentSpread<=0;
  document.getElementById('btn-next').disabled = currentSpread>=totalSpreads-1;
}}

document.getElementById('btn-prev').addEventListener('click',()=>flipTo(-1));
document.getElementById('btn-next').addEventListener('click',()=>flipTo(1));
document.addEventListener('keydown',e=>{{
  if(e.key==='ArrowRight'||e.key==='ArrowDown') flipTo(1);
  if(e.key==='ArrowLeft' ||e.key==='ArrowUp')   flipTo(-1);
}});
let tx=null;
document.addEventListener('touchstart',e=>{{tx=e.touches[0].clientX;}});
document.addEventListener('touchend',e=>{{
  if(tx===null)return;
  const dx=e.changedTouches[0].clientX-tx;
  if(Math.abs(dx)>50) flipTo(dx<0?1:-1);
  tx=null;
}});
document.getElementById('stage').addEventListener('click',e=>{{
  if(e.target.closest('.nav-btn')) return;
  const w=window.innerWidth;
  if(e.clientX<w*0.25) flipTo(-1);
  else if(e.clientX>w*0.75) flipTo(1);
}});
document.addEventListener('click',()=>{{try{{getAudio();}}catch(e){{}}}},{{once:true}});
if(isMobile){{
  const pw=parseInt(getComputedStyle(document.documentElement).getPropertyValue('--pw'));
  document.getElementById('book-wrap').style.width=pw+'px';
}}
renderSpread(0);
updateUI();
setTimeout(()=>document.getElementById('loading').classList.add('hidden'),350);
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ index.html generado ({len(pages_b64)} páginas)")

# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    base_dir    = Path(__file__).parent
    output_path = base_dir / "index.html"

    title    = os.environ.get("BOOK_TITLE",    "Libro Digital")
    subtitle = os.environ.get("BOOK_SUBTITLE", "")
    file_id  = os.environ.get("DRIVE_FILE_ID")

    if "--demo" in sys.argv or not file_id:
        print("ℹ️  Modo demo — crea un PDF de prueba para ver el libro funcionando.")
        print("   Para usar tu documento real, configura DRIVE_FILE_ID en GitHub Secrets.")
        # Demo: crear un PDF mínimo con texto
        try:
            from reportlab.pdfgen import canvas as rl_canvas
            from reportlab.lib.pagesizes import A4
            buf = io.BytesIO()
            c = rl_canvas.Canvas(buf, pagesize=A4)
            c.setFont("Helvetica-Bold", 24)
            c.drawString(100, 750, title)
            c.setFont("Helvetica", 16)
            c.drawString(100, 710, subtitle)
            c.setFont("Helvetica", 12)
            c.drawString(100, 650, "Página de ejemplo — conecta tu documento Word en Drive")
            c.drawString(100, 630, "para ver el contenido real aquí.")
            c.showPage()
            c.setFont("Helvetica-Bold", 18)
            c.drawString(100, 750, "Segunda página de ejemplo")
            c.setFont("Helvetica", 12)
            c.drawString(100, 710, "El contenido de tu Word aparecerá aquí.")
            c.save()
            pdf_bytes = buf.getvalue()
        except ImportError:
            print("   (reportlab no instalado, usando PDF mínimo)")
            # PDF mínimo hardcodeado
            pdf_bytes = b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\nxref\n0 4\n0000000000 65535 f\n0000000009 00000 n\n0000000058 00000 n\n0000000115 00000 n\ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF"
    else:
        print(f"📥 Descargando documento desde Drive (ID: {file_id})...")
        pdf_bytes = download_as_pdf(file_id)

    print("🖼️  Convirtiendo PDF a imágenes...")
    pages_b64 = pdf_to_images_b64(pdf_bytes, dpi=DPI)

    print("📖 Generando libro digital...")
    generate_html(pages_b64, title, subtitle, str(output_path))

if __name__ == "__main__":
    main()

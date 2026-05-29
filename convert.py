#!/usr/bin/env python3
"""
convert.py — Descarga el .docx desde Google Drive, lo convierte a PDF
             con LibreOffice, y genera el libro digital con las páginas
             como imágenes (formato 100% preservado).

Dependencias Python:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pdf2image pillow

Dependencias sistema (instaladas en publish.yml):
    sudo apt-get install -y libreoffice poppler-utils

Variables de entorno (GitHub Secrets):
    GOOGLE_SERVICE_ACCOUNT_JSON
    DRIVE_FILE_ID
    BOOK_TITLE    (opcional)
    BOOK_SUBTITLE (opcional)
"""

import json, os, sys, io, base64, subprocess, tempfile, shutil
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

DPI = 150  # calidad de imagen (subir a 200 para más nitidez)

# ──────────────────────────────────────────────
# 1. DESCARGA EL .docx DESDE GOOGLE DRIVE
# ──────────────────────────────────────────────

def download_docx(file_id: str, dest_path: str) -> None:
    if not HAS_GOOGLE:
        raise RuntimeError("Falta google-api-python-client.")
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not sa_json:
        raise ValueError("Variable GOOGLE_SERVICE_ACCOUNT_JSON no encontrada.")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json),
        scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    meta = service.files().get(fileId=file_id, fields="name,mimeType").execute()
    print(f"📄 Archivo: {meta.get('name')} ({meta.get('mimeType')})")

    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        if status:
            print(f"   Descargando... {int(status.progress() * 100)}%")
    with open(dest_path, "wb") as f:
        f.write(fh.getvalue())
    print(f"✅ Descargado: {dest_path} ({Path(dest_path).stat().st_size // 1024} KB)")

# ──────────────────────────────────────────────
# 2. CONVERTIR .docx → PDF CON LIBREOFFICE
# ──────────────────────────────────────────────

def docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """
    Usa LibreOffice headless para convertir docx a PDF.
    Retorna la ruta del PDF generado.
    """
    print("📄 Convirtiendo a PDF con LibreOffice...")
    result = subprocess.run(
        [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", output_dir, docx_path
        ],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)
        raise RuntimeError(f"LibreOffice falló con código {result.returncode}")

    # El PDF se genera con el mismo nombre base
    docx_name = Path(docx_path).stem
    pdf_path  = Path(output_dir) / f"{docx_name}.pdf"
    if not pdf_path.exists():
        # Buscar cualquier PDF generado
        pdfs = list(Path(output_dir).glob("*.pdf"))
        if not pdfs:
            raise RuntimeError("LibreOffice no generó ningún PDF")
        pdf_path = pdfs[0]

    print(f"✅ PDF generado: {pdf_path} ({pdf_path.stat().st_size // 1024} KB)")
    return str(pdf_path)

# ──────────────────────────────────────────────
# 3. PDF → IMÁGENES BASE64
# ──────────────────────────────────────────────

def pdf_to_images_b64(pdf_path: str, dpi: int = DPI) -> list:
    if not HAS_PDF2IMAGE:
        raise RuntimeError("Falta pdf2image. Ejecuta: pip install pdf2image pillow")
    print(f"🖼️  Convirtiendo páginas a imágenes (DPI={dpi})...")
    images = convert_from_path(pdf_path, dpi=dpi, fmt="png")
    print(f"   {len(images)} páginas encontradas")
    b64_pages = []
    for i, img in enumerate(images):
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        b64_pages.append(b64)
        print(f"   Página {i+1}/{len(images)} ({len(b64)//1024} KB)")
    return b64_pages

# ──────────────────────────────────────────────
# 4. GENERAR index.html
# ──────────────────────────────────────────────

def generate_html(pages_b64: list, title: str, subtitle: str, output_path: str):
    if not pages_b64:
        raise RuntimeError("No hay páginas para generar")

    # PAGES[0]=pág1, PAGES[1]=pág2, PAGES[2]=pág3...
    # Spread N: left=PAGES[N*2], right=PAGES[N*2+1]
    parts = []
    for b64 in pages_b64:
        parts.append('{"type":"img","src":"data:image/png;base64,' + b64 + '"}')
    if len(parts) % 2 != 0:
        parts.append('{"type":"blank"}')
    pages_json = "[" + ",".join(parts) + "]"

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
  --gold:#B8862A;--gold2:#D4A843;--gold3:#F0C96A;--muted:#8A6F52;--cream:#F7F0E0;
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
  position:relative;width:calc(var(--pw)*2);height:var(--ph);
  transform-style:preserve-3d;transform:rotateX(2deg);
  filter:drop-shadow(0 32px 64px rgba(0,0,0,.8)) drop-shadow(0 8px 16px rgba(0,0,0,.5));
}}
#book-wrap::after{{
  content:'';position:absolute;left:50%;top:0;bottom:0;width:5px;margin-left:-2px;
  background:linear-gradient(to right,rgba(10,5,2,.8),rgba(80,45,20,.5),rgba(10,5,2,.8));
  z-index:50;pointer-events:none;
}}
.page{{
  position:absolute;width:var(--pw);height:var(--ph);top:0;
  background:var(--cream);backface-visibility:hidden;overflow:hidden;
}}
.page.left-page{{left:0;transform-origin:right center;}}
.page.right-page{{left:var(--pw);transform-origin:left center;}}
.page.left-page::after{{
  content:'';position:absolute;top:0;right:0;width:28px;height:100%;
  background:linear-gradient(to left,rgba(20,10,4,.15),transparent);z-index:2;pointer-events:none;
}}
.page.right-page::before{{
  content:'';position:absolute;top:0;left:0;width:28px;height:100%;
  background:linear-gradient(to right,rgba(20,10,4,.1),transparent);z-index:2;pointer-events:none;
}}
.page img{{width:100%;height:100%;object-fit:contain;object-position:center top;display:block;user-select:none;}}
.page.blank-page{{background:var(--cream);}}
#flip-layer{{
  position:absolute;width:var(--pw);height:var(--ph);top:0;
  transform-style:preserve-3d;pointer-events:none;z-index:100;
}}
#flip-front,#flip-back{{position:absolute;inset:0;backface-visibility:hidden;overflow:hidden;}}
#flip-back{{transform:rotateY(180deg);}}
#flip-shadow{{position:absolute;inset:0;z-index:10;pointer-events:none;opacity:0;}}
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
.nav-btn:hover:not(:disabled){{background:rgba(61,43,31,.95);border-color:var(--gold3);color:var(--gold3);}}
.nav-btn:active:not(:disabled){{transform:scale(.93);}}
.nav-btn:disabled{{opacity:.25;cursor:not-allowed;}}
.nav-btn .hint{{position:absolute;bottom:-20px;font-size:10px;color:var(--muted);font-family:'Lora',serif;font-style:italic;white-space:nowrap;}}
#page-info{{font-size:13px;color:var(--gold2);font-family:'Playfair Display',serif;min-width:80px;text-align:center;}}
#loading{{
  position:fixed;inset:0;z-index:300;background:#1A0F08;
  display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;
  transition:opacity .6s;
}}
#loading.hidden{{opacity:0;pointer-events:none;}}
.spinner{{width:36px;height:36px;border:2px solid rgba(184,134,42,.2);border-top-color:var(--gold);border-radius:50%;animation:spin .85s linear infinite;}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
#loading p{{font-family:'Playfair Display',serif;color:var(--gold2);font-style:italic;font-size:15px;}}
@media(max-width:980px){{
  :root{{--pw:300px;--ph:420px;}}
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
const PAGES={pages_json};
let audioCtx=null;
function getAudio(){{if(!audioCtx)audioCtx=new(window.AudioContext||window.webkitAudioContext)();return audioCtx;}}
function playFlip(){{
  try{{
    const ctx=getAudio(),sr=ctx.sampleRate;
    const buf=ctx.createBuffer(1,sr*.22,sr),d=buf.getChannelData(0);
    for(let i=0;i<d.length;i++){{
      const t=i/sr,s=Math.random()*2-1;
      const env=Math.exp(-t*18)*(1-Math.exp(-t*180));
      const thump=Math.exp(-t*40)*Math.sin(2*Math.PI*80*t)*.35;
      const crk=s*Math.sin(2*Math.PI*(800+400*Math.sin(t*60))*t)*.5;
      d[i]=(s*env*.4+crk*env*.4+thump)*.55;
    }}
    const src=ctx.createBufferSource();src.buffer=buf;
    const bpf=ctx.createBiquadFilter();bpf.type='bandpass';bpf.frequency.value=1800;bpf.Q.value=.6;
    const hpf=ctx.createBiquadFilter();hpf.type='highpass';hpf.frequency.value=300;
    const g=ctx.createGain();g.gain.value=.55;
    src.connect(hpf);hpf.connect(bpf);bpf.connect(g);g.connect(ctx.destination);src.start();
  }}catch(e){{}}
}}
function mkPage(data,side){{
  const d=document.createElement('div');
  d.className='page '+side;
  if(!data||data.type==='blank'){{d.classList.add('blank-page');return d;}}
  const img=document.createElement('img');
  img.src=data.src;img.alt='';img.draggable=false;
  d.appendChild(img);return d;
}}
const isMobile=window.innerWidth<=980;
const totalSpreads=Math.ceil(PAGES.length/2);
let cur=0,busy=false;
const RS=document.getElementById('right-slot');
const LS=document.getElementById('left-slot');
const FL=document.getElementById('flip-layer');
const FF=document.getElementById('flip-front');
const FB=document.getElementById('flip-back');
const FS=document.getElementById('flip-shadow');
function sp(n){{return{{l:PAGES[n*2]||null,r:PAGES[n*2+1]||null}};}}
function render(n){{
  const s=sp(n);
  RS.innerHTML='';LS.innerHTML='';
  if(!isMobile)LS.appendChild(mkPage(s.l,'left-page'));
  RS.appendChild(mkPage(s.r,'right-page'));
}}
const MS=480;
function go(dir){{
  if(busy)return;
  const nx=cur+dir;
  if(nx<0||nx>=totalSpreads)return;
  busy=true;playFlip();
  const cs=sp(cur),ns=sp(nx);
  if(isMobile){{
    const f=RS.firstChild;
    if(f){{f.style.transition='opacity .25s';f.style.opacity='0';}}
    setTimeout(()=>{{
      cur=nx;render(nx);
      const t=RS.firstChild;
      if(t){{t.style.opacity='0';t.style.transition='opacity .25s';}}
      requestAnimationFrame(()=>{{if(t)t.style.opacity='1';}});
      ui();busy=false;
    }},250);return;
  }}
  if(dir>0){{
    FL.style.left='var(--pw)';FL.style.transformOrigin='left center';
    FF.innerHTML='';FS.style.background='linear-gradient(to left,rgba(0,0,0,.38),transparent 72%)';FF.appendChild(FS);
    if(cs.r){{const e=mkPage(cs.r,'right-page');e.style.cssText='position:absolute;inset:0';FF.appendChild(e);}}
    FB.innerHTML='';
    if(ns.l){{const e=mkPage(ns.l,'left-page');e.style.cssText='position:absolute;inset:0;transform:scaleX(-1)';FB.appendChild(e);}}
    RS.innerHTML='';RS.appendChild(mkPage(ns.r,'right-page'));
  }}else{{
    FL.style.left='0';FL.style.transformOrigin='right center';
    FF.innerHTML='';FS.style.background='linear-gradient(to right,rgba(0,0,0,.38),transparent 72%)';FF.appendChild(FS);
    if(cs.l){{const e=mkPage(cs.l,'left-page');e.style.cssText='position:absolute;inset:0';FF.appendChild(e);}}
    FB.innerHTML='';
    if(ns.r){{const e=mkPage(ns.r,'right-page');e.style.cssText='position:absolute;inset:0;transform:scaleX(-1)';FB.appendChild(e);}}
    LS.innerHTML='';LS.appendChild(mkPage(ns.l,'left-page'));
  }}
  FL.style.transition='none';FL.style.transform='rotateY(0deg)';FS.style.opacity='0';
  const deg=dir>0?-180:180;
  requestAnimationFrame(()=>requestAnimationFrame(()=>{{
    FL.style.transition=`transform ${{MS}}ms cubic-bezier(.645,.045,.355,1)`;
    FS.style.transition=`opacity ${{MS}}ms`;
    FL.style.transform=`rotateY(${{deg}}deg)`;FS.style.opacity='1';
  }}));
  setTimeout(()=>{{
    FL.style.transition='none';FL.style.transform='rotateY(0deg)';FS.style.opacity='0';
    cur=nx;render(nx);ui();busy=false;
  }},MS+30);
}}
function ui(){{
  const p=cur*2+1;
  document.getElementById('page-counter').textContent=`Pág. ${{p}}`;
  document.getElementById('page-info').textContent=`${{cur+1}} / ${{totalSpreads}}`;
  document.getElementById('btn-prev').disabled=cur<=0;
  document.getElementById('btn-next').disabled=cur>=totalSpreads-1;
}}
document.getElementById('btn-prev').addEventListener('click',()=>go(-1));
document.getElementById('btn-next').addEventListener('click',()=>go(1));
document.addEventListener('keydown',e=>{{
  if(e.key==='ArrowRight'||e.key==='ArrowDown')go(1);
  if(e.key==='ArrowLeft'||e.key==='ArrowUp')go(-1);
}});
let tx=null;
document.addEventListener('touchstart',e=>{{tx=e.touches[0].clientX;}});
document.addEventListener('touchend',e=>{{
  if(tx===null)return;const dx=e.changedTouches[0].clientX-tx;
  if(Math.abs(dx)>50)go(dx<0?1:-1);tx=null;
}});
document.getElementById('stage').addEventListener('click',e=>{{
  if(e.target.closest('.nav-btn'))return;
  const w=window.innerWidth;
  if(e.clientX<w*.25)go(-1);else if(e.clientX>w*.75)go(1);
}});
document.addEventListener('click',()=>{{try{{getAudio();}}catch(e){{}}}},{{once:true}});
if(isMobile){{
  document.getElementById('book-wrap').style.width=
    getComputedStyle(document.documentElement).getPropertyValue('--pw');
}}
render(0);ui();
setTimeout(()=>document.getElementById('loading').classList.add('hidden'),350);
</script>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ index.html generado ({len(pages_b64)} páginas, {Path(output_path).stat().st_size//1024} KB)")

# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    base_dir    = Path(__file__).parent
    output_path = base_dir / "index.html"
    title       = os.environ.get("BOOK_TITLE",    "Libro Digital")
    subtitle    = os.environ.get("BOOK_SUBTITLE",  "")
    file_id     = os.environ.get("DRIVE_FILE_ID")

    if "--demo" in sys.argv or not file_id:
        print("ℹ️  Sin DRIVE_FILE_ID — generando página de ejemplo.")
        # Página demo simple sin PDF
        demo_b64 = generate_demo_page(title, subtitle)
        generate_html([demo_b64], title, subtitle, str(output_path))
        return

    tmp_dir = tempfile.mkdtemp()
    try:
        # 1. Descargar .docx
        docx_path = os.path.join(tmp_dir, "documento.docx")
        download_docx(file_id, docx_path)

        # 2. Convertir a PDF con LibreOffice
        pdf_path = docx_to_pdf(docx_path, tmp_dir)

        # 3. PDF → imágenes
        pages_b64 = pdf_to_images_b64(pdf_path, dpi=DPI)

        # 4. Generar HTML
        generate_html(pages_b64, title, subtitle, str(output_path))

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def generate_demo_page(title, subtitle):
    """Genera una imagen PNG de demo en base64 sin dependencias externas."""
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("RGB", (960, 1320), color=(247, 240, 224))
        draw = ImageDraw.Draw(img)
        draw.rectangle([40, 40, 920, 1280], outline=(184, 134, 42), width=2)
        draw.text((480, 400), title,    fill=(61, 43, 31), anchor="mm")
        draw.text((480, 460), subtitle, fill=(138, 111, 82), anchor="mm")
        draw.text((480, 560), "Conecta tu documento Word en Drive", fill=(138,111,82), anchor="mm")
        draw.text((480, 600), "para ver el contenido real aquí.",   fill=(138,111,82), anchor="mm")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode()
    except Exception:
        # Mínimo: PNG blanco 1x1
        return "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwADhQGAWjR9awAAAABJRU5ErkJggg=="


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
app.py – Flask web app for the AI-powered CODIFY Converter
"""
from __future__ import annotations
import io, json, re, traceback
from datetime import datetime
from flask import Flask, request, send_file, render_template_string, jsonify, Response, stream_with_context
from openpyxl import load_workbook
from codify import convert

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024 * 1024

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>CI Benchmarking → CODIFY Converter</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --bg: #f0f2f5; --surface: #fff; --surface-2: #f7f8fa;
    --primary: #1a4f8a; --primary-light: #e8f1fb; --primary-hover: #1560a8;
    --text: #111827; --text-muted: #6b7280; --text-soft: #9ca3af;
    --border: #e5e7eb; --border-strong: #d1d5db;
    --success: #065f46; --success-bg: #ecfdf5; --success-border: #6ee7b7;
    --error: #991b1b; --error-bg: #fff1f2; --error-border: #fca5a5;
    --radius: 10px; --radius-sm: 6px;
    --shadow: 0 1px 4px rgba(0,0,0,.06),0 4px 16px rgba(0,0,0,.04);
  }
  body { font-family: -apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
    background:var(--bg); color:var(--text); min-height:100vh;
    padding:48px 20px 64px; line-height:1.5; -webkit-font-smoothing:antialiased; }
  .page { max-width:640px; margin:0 auto; }

  .header { display:flex; align-items:center; gap:14px; margin-bottom:32px; }
  .header-icon { width:44px;height:44px;background:var(--primary);border-radius:var(--radius-sm);
    display:flex;align-items:center;justify-content:center;flex-shrink:0; }
  .header-icon svg { width:22px;height:22px;stroke:#fff; }
  .header-text h1 { font-size:20px;font-weight:700;letter-spacing:-.3px; }
  .header-text p { font-size:13.5px;color:var(--text-muted);margin-top:2px; }

  .card { background:var(--surface);border:1px solid var(--border);
    border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden; }
  .card-body { padding:28px; }

  .dropzone { border:2px dashed var(--border-strong);border-radius:var(--radius);
    padding:40px 24px;text-align:center;cursor:pointer;
    transition:border-color .15s,background .15s;background:var(--surface-2);position:relative; }
  .dropzone:hover,.dropzone.over { border-color:var(--primary);background:var(--primary-light); }
  .dropzone.has-file { border-style:solid;border-color:var(--primary);
    background:var(--primary-light);padding:20px 24px; }
  .dropzone input[type=file] { position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%; }
  .dz-idle { display:flex;flex-direction:column;align-items:center;gap:10px; }
  .dz-icon { width:48px;height:48px;background:#dbeafe;border-radius:50%;
    display:flex;align-items:center;justify-content:center; }
  .dz-icon svg { width:22px;height:22px;stroke:var(--primary); }
  .dz-title { font-size:15px;font-weight:600; }
  .dz-sub { font-size:13px;color:var(--text-muted); }
  .dz-badge { font-size:11px;font-weight:600;letter-spacing:.4px;padding:3px 8px;
    border-radius:4px;background:#e0e7ef;color:var(--primary);margin-top:4px; }
  .dz-file { display:none;align-items:center;gap:12px; }
  .dz-file.show { display:flex; }
  .dz-file-icon { width:40px;height:40px;flex-shrink:0;background:var(--primary);
    border-radius:8px;display:flex;align-items:center;justify-content:center; }
  .dz-file-icon svg { width:20px;height:20px;stroke:#fff; }
  .dz-file-info { flex:1;text-align:left;min-width:0; }
  .dz-file-name { font-size:14px;font-weight:600;white-space:nowrap;
    overflow:hidden;text-overflow:ellipsis; }
  .dz-file-size { font-size:12px;color:var(--text-muted);margin-top:2px; }
  .dz-remove { background:none;border:none;cursor:pointer;padding:6px;border-radius:6px;
    color:var(--text-muted);transition:background .1s,color .1s;
    z-index:1;position:relative;flex-shrink:0; }
  .dz-remove:hover { background:#fee2e2;color:#dc2626; }
  .dz-remove svg { width:16px;height:16px;stroke:currentColor;display:block; }

  .btn-convert { margin-top:16px;width:100%;padding:13px 20px;background:var(--primary);
    color:#fff;font-size:15px;font-weight:600;border:none;border-radius:var(--radius-sm);
    cursor:pointer;transition:background .15s,opacity .15s,transform .1s;
    display:flex;align-items:center;justify-content:center;gap:8px; }
  .btn-convert:hover:not(:disabled) { background:var(--primary-hover); }
  .btn-convert:active:not(:disabled) { transform:scale(.99); }
  .btn-convert:disabled { opacity:.45;cursor:not-allowed; }
  .btn-convert svg { width:18px;height:18px;stroke:#fff;flex-shrink:0; }

  /* Progress log */
  .progress-box { display:none;margin-top:16px;border:1px solid var(--border);
    border-radius:var(--radius-sm);background:#0f172a;overflow:hidden; }
  .progress-box.show { display:block; }
  .progress-header { padding:8px 14px;background:#1e293b;
    font-size:11px;font-weight:700;color:#94a3b8;letter-spacing:.5px;text-transform:uppercase; }
  .progress-log { padding:12px 14px;min-height:60px;max-height:180px;overflow-y:auto;
    font-family:'SFMono-Regular',Consolas,monospace;font-size:12px;
    color:#e2e8f0;line-height:1.7; }
  .progress-log .line-ok   { color:#4ade80; }
  .progress-log .line-wait { color:#facc15; }
  .progress-log .line-err  { color:#f87171; }

  .status { display:none;margin-top:16px;padding:14px 16px;border-radius:var(--radius-sm);
    font-size:13.5px;line-height:1.55;border:1px solid transparent; }
  .status.show { display:block; }
  .status.ok { background:var(--success-bg);color:var(--success);border-color:var(--success-border); }
  .status.err { background:var(--error-bg);color:var(--error);border-color:var(--error-border); }

  .summary-pills { display:flex;flex-wrap:wrap;gap:8px;margin-top:10px; }
  .pill { display:flex;align-items:center;gap:5px;background:#d1fae5;color:#065f46;
    border-radius:20px;padding:4px 12px;font-size:12.5px;font-weight:600; }
  .new-sheets { margin-top:8px;font-size:12.5px;color:var(--success); }

  .spinner { width:16px;height:16px;border:2px solid rgba(255,255,255,.35);
    border-top-color:#fff;border-radius:50%;animation:spin .65s linear infinite;flex-shrink:0; }
  @keyframes spin { to{transform:rotate(360deg)} }

  .divider { height:1px;background:var(--border); }
  .info { padding:20px 28px 24px;background:var(--surface-2); }
  .info-title { font-size:12px;font-weight:700;letter-spacing:.6px;
    text-transform:uppercase;color:var(--text-soft);margin-bottom:12px; }
  .info-list { list-style:none;display:flex;flex-direction:column;gap:8px; }
  .info-list li { display:flex;align-items:flex-start;gap:10px;
    font-size:13.5px;color:var(--text-muted); }
  .info-list li svg { width:15px;height:15px;stroke:var(--primary);flex-shrink:0;margin-top:1px; }
  code { font-family:'SFMono-Regular',Consolas,monospace;font-size:12px;
    background:#e5e7eb;padding:1px 5px;border-radius:3px;color:#374151; }
  .ai-badge { display:inline-flex;align-items:center;gap:4px;
    background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;
    font-size:11px;font-weight:700;padding:2px 8px;border-radius:20px;
    letter-spacing:.3px;vertical-align:middle;margin-left:6px; }
</style>
</head>
<body>
<div class="page">
  <div class="header">
    <div class="header-icon">
      <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
        <line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>
      </svg>
    </div>
    <div class="header-text">
      <h1>CI Benchmarking → CODIFY <span class="ai-badge">✦ Free Local AI</span></h1>
      <p>Upload a benchmarking workbook and download a fully structured CODIFY comparison file</p>
    </div>
  </div>

  <div class="card">
    <div class="card-body">
      <div id="drop" class="dropzone" role="button">
        <input type="file" id="file" accept=".xlsx,.xlsm">
        <div class="dz-idle" id="dz-idle">
          <div class="dz-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
              <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
          </div>
          <div><div class="dz-title">Drop your benchmarking workbook</div>
               <div class="dz-sub">or click to browse</div></div>
          <span class="dz-badge">.xlsx &nbsp;/&nbsp; .xlsm</span>
        </div>
        <div class="dz-file" id="dz-file">
          <div class="dz-file-icon">
            <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
              <polyline points="14 2 14 8 20 8"/>
            </svg>
          </div>
          <div class="dz-file-info">
            <div class="dz-file-name" id="dz-file-name">—</div>
            <div class="dz-file-size" id="dz-file-size">—</div>
          </div>
          <button class="dz-remove" id="dz-remove" type="button" title="Remove">
            <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
              <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          </button>
        </div>
      </div>

      <button id="convert" class="btn-convert" disabled type="button">
        <svg id="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2"
             stroke-linecap="round" stroke-linejoin="round">
          <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
        </svg>
        <span id="btn-label">Convert with Free Local AI</span>
      </button>

      <!-- Live progress log -->
      <div class="progress-box" id="progress-box">
        <div class="progress-header">AI Processing Log</div>
        <div class="progress-log" id="progress-log"></div>
      </div>

      <div id="status" class="status" role="status" aria-live="polite"></div>
    </div>

    <div class="divider"></div>
    <div class="info">
      <div class="info-title">How it works</div>
      <ul class="info-list">
        <li>
          <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
          Free local AI reads your benchmarking sheet and maps every feature to the right CODIFY sheet automatically
        </li>
        <li>
          <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
          New features not in the standard schema automatically get their own generated sheet
        </li>
        <li>
          <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
          Every sheet includes a <code>raw_text</code> column with the original source value for verification
        </li>
        <li>
          <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
          Works with any layout, language, or column structure — no configuration needed
        </li>
      </ul>
    </div>
  </div>
</div>

<script>
const dropEl   = document.getElementById('drop');
const fileInput= document.getElementById('file');
const dzIdle   = document.getElementById('dz-idle');
const dzFile   = document.getElementById('dz-file');
const dzName   = document.getElementById('dz-file-name');
const dzSize   = document.getElementById('dz-file-size');
const dzRemove = document.getElementById('dz-remove');
const convertBtn=document.getElementById('convert');
const btnLabel = document.getElementById('btn-label');
const statusBox= document.getElementById('status');
const progBox  = document.getElementById('progress-box');
const progLog  = document.getElementById('progress-log');
let selectedFile = null, downloadUrl = null;

function fmtSize(b){
  if(b<1024)return b+' B';
  if(b<1048576)return(b/1024).toFixed(1)+' KB';
  return(b/1048576).toFixed(1)+' MB';
}
function setFile(f){
  selectedFile=f;
  if(f){
    dzName.textContent=f.name; dzSize.textContent=fmtSize(f.size);
    dzIdle.style.display='none'; dzFile.classList.add('show');
    dropEl.classList.add('has-file'); convertBtn.disabled=false;
    clearStatus(); clearLog();
  } else {
    dzIdle.style.display=''; dzFile.classList.remove('show');
    dropEl.classList.remove('has-file'); convertBtn.disabled=true;
    fileInput.value=''; clearStatus(); clearLog();
  }
}
function clearStatus(){ statusBox.className='status'; statusBox.innerHTML=''; }
function clearLog(){ progBox.classList.remove('show'); progLog.innerHTML=''; }
function addLog(msg, cls='line-wait'){
  const line=document.createElement('div');
  line.className=cls; line.textContent='› '+msg;
  progLog.appendChild(line); progLog.scrollTop=progLog.scrollHeight;
}
function showStatus(html,type){ statusBox.innerHTML=html; statusBox.className='status show '+type; }
function setBusy(busy){
  convertBtn.disabled=busy;
  if(busy){
    document.getElementById('btn-icon').outerHTML='<div class="spinner" id="btn-icon"></div>';
    btnLabel.textContent='Processing…';
  } else {
    const el=document.getElementById('btn-icon');
    if(el) el.outerHTML='<svg id="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>';
    btnLabel.textContent='Convert with Free Local AI';
  }
}

dzRemove.addEventListener('click',e=>{e.stopPropagation();setFile(null);});
fileInput.addEventListener('change',e=>{const f=e.target.files[0];if(f)setFile(f);});
['dragenter','dragover'].forEach(ev=>dropEl.addEventListener(ev,e=>{e.preventDefault();if(!selectedFile)dropEl.classList.add('over');}));
['dragleave','drop'].forEach(ev=>dropEl.addEventListener(ev,e=>{e.preventDefault();dropEl.classList.remove('over');}));
dropEl.addEventListener('drop',e=>{const f=e.dataTransfer.files[0];if(f)setFile(f);});

convertBtn.addEventListener('click', async ()=>{
  if(!selectedFile) return;
  setBusy(true); clearLog();
  progBox.classList.add('show');
  addLog('Uploading file…');

  const fd=new FormData();
  fd.append('file', selectedFile);

  try {
    // Stream progress via SSE, then get the file
    const r = await fetch('/convert', {method:'POST', body:fd});
    if(!r.ok){
      const err=await r.json().catch(()=>({error:'Server error '+r.status}));
      addLog('Failed: '+(err.error||'Unknown error'), 'line-err');
      showStatus('<strong>Error:</strong> '+(err.error||'Conversion failed'), 'err');
      setBusy(false); return;
    }

    const meta = JSON.parse(r.headers.get('X-Codify-Summary')||'null');
    // Log progress messages embedded in headers
    const steps = (r.headers.get('X-Codify-Steps')||'').split('|').filter(Boolean);
    steps.forEach(s=>addLog(s,'line-ok'));
    addLog('Building workbook…','line-ok');

    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href=url;
    a.download=selectedFile.name.replace(/\.(xlsx|xlsm)$/i,'')+('_Codified.xlsx');
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);

    addLog('Done — download started.','line-ok');

    let html='<strong>Conversion complete — download started.</strong>';
    if(meta){
      html+='<div class="summary-pills">'
        +'<span class="pill">'+meta.companies+' companies</span>'
        +'<span class="pill">'+meta.products+' products</span>'
        +'<span class="pill">'+meta.features+' features</span>'
        +'<span class="pill">'+meta.sheets+' sheets</span>'
        +'</div>';
      if(meta.new_sheets && meta.new_sheets.length){
        html+='<div class="new-sheets">✦ '+meta.new_sheets.length
          +' new sheet(s) auto-generated: '+meta.new_sheets.join(', ')+'</div>';
      }
    }
    showStatus(html,'ok');
  } catch(err){
    addLog('Network error: '+err.message,'line-err');
    showStatus('<strong>Error:</strong> '+err.message,'err');
  } finally { setBusy(false); }
});
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(INDEX_HTML)


@app.route('/convert', methods=['POST'])
def do_convert():
    f = request.files.get('file')
    if not f or not f.filename:
        return jsonify({'error': 'No file uploaded.'}), 400
    if not f.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': 'Please upload an .xlsx or .xlsm file.'}), 400

    try:
        wb_in = load_workbook(io.BytesIO(f.read()), data_only=True)
    except Exception as e:
        return jsonify({'error': f'Could not read workbook: {e}'}), 400

    steps = []
    def progress(msg):
        steps.append(msg)
        app.logger.info(msg)

    try:
        wb_out, summary = convert(wb_in, progress_cb=progress)
    except Exception as e:
        app.logger.error(traceback.format_exc())
        return jsonify({'error': f'Conversion failed: {e}'}), 500

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)

    download_name = re.sub(r'\.(xlsx|xlsm)$', '', f.filename, flags=re.IGNORECASE) + '_Codified.xlsx'
    response = send_file(
        buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=download_name,
    )
    response.headers['X-Codify-Summary'] = json.dumps(summary)
    response.headers['X-Codify-Steps']   = '|'.join(steps)
    response.headers['Access-Control-Expose-Headers'] = 'X-Codify-Summary,X-Codify-Steps'
    return response


@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'timestamp': datetime.utcnow().isoformat()})


if __name__ == '__main__':
    print('━' * 52)
    print('  CI Benchmarking → CODIFY Converter  [Free local AI / Ollama]')
    print('  Open: http://127.0.0.1:5000')
    print('  Start Ollama and pull a model: ollama pull llama3.1:8b')
    print('  Press Ctrl+C to stop')
    print('━' * 52)
    app.run(host='127.0.0.1', port=5000, debug=False)

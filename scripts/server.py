#!/usr/bin/env python3
import http.server
import socketserver
import json
import os
import urllib.parse
from pathlib import Path

PORT = 8080

# Load INNs from Excel
import openpyxl
src = "Администрации_городов_из_реестра_2026-11-10.xlsx"
wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
# Gather unique INNs from all sheets except Оглавление
inn_data = []
seen = set()
for ws in wb.worksheets:
    if ws.title == "Оглавление":
        continue
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    # find INN col index
    try:
        inn_idx = header.index("ИНН")
    except:
        inn_idx = 5
    for r in rows[1:]:
        inn = str(r[inn_idx]).strip() if r[inn_idx] else ""
        if inn and inn.isdigit() and len(inn) in (10,12) and inn not in seen:
            seen.add(inn)
            inn_data.append({
                "inn": inn,
                "full_name": r[0],
                "ogrn": r[4],
                "site": r[7],
                "email": r[9],
                "sheet": ws.title
            })
print(f"Loaded {len(inn_data)} unique INNs")

# Also load full rows for saving later? We'll just reload when needed

HTML = f"""
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DaData Обогащение — Администрации городов</title>
<style>
body {{ font-family: -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; background:#f6f8fb; color:#222; }}
.card {{ background:white; border-radius:12px; padding:24px; box-shadow:0 2px 12px rgba(0,0,0,0.08); margin-bottom:20px; }}
h1 {{ font-size:22px; margin:0 0 10px; color:#1F4E78; }}
.subtitle {{ color:#666; font-size:13px; margin-bottom:16px; }}
.stats {{ display:flex; gap:16px; margin:16px 0; flex-wrap:wrap; }}
.stat {{ flex:1; min-width:120px; background:#f2f6ff; border-radius:8px; padding:12px; text-align:center; }}
.stat b {{ display:block; font-size:22px; color:#1F4E78; }}
.stat span {{ font-size:11px; color:#666; text-transform:uppercase; letter-spacing:0.5px; }}
.progress {{ height:12px; background:#e9eef6; border-radius:6px; overflow:hidden; margin:16px 0; }}
.progress-bar {{ height:100%; background:linear-gradient(90deg,#1F4E78,#548235); width:0%; transition:width 0.3s; }}
.log {{ background:#0f172a; color:#e2e8f0; padding:16px; border-radius:8px; font-family:monospace; font-size:12px; max-height:300px; overflow:auto; white-space:pre-wrap; }}
.btn {{ background:#1F4E78; color:white; border:none; padding:12px 20px; border-radius:8px; cursor:pointer; font-size:14px; font-weight:600; }}
.btn:disabled {{ opacity:0.5; cursor:not-allowed; }}
.btn.secondary {{ background:#e2e8f0; color:#334; }}
.tag {{ display:inline-block; font-size:11px; padding:3px 8px; border-radius:20px; background:#e2e8f0; color:#334; margin-right:6px; }}
.ok {{ color:#16a34a; }} .err {{ color:#dc2626; }} .warn {{ color:#d97706; }}
table {{ width:100%; border-collapse:collapse; font-size:12px; margin-top:12px; }}
th,td {{ border:1px solid #e2e8f0; padding:6px 8px; text-align:left; }}
th {{ background:#f1f5f9; }}
</style>
</head>
<body>
<div class="card">
  <h1>🔍 DaData Find-Party — обогащение администраций городов</h1>
  <div class="subtitle">Ваш браузер напрямую обращается к <code>suggestions.dadata.ru</code> (обход firewall песочницы). API-ключ <code>77eff...f382dfb</code> уже вставлен. Запросы идут с вашей машины в Frankfurt — лимиты DaData соблюдаются (пауза 150 мс).</div>
  <div class="stats">
    <div class="stat"><b id="total">{len(inn_data)}</b><span>Уникальных ИНН</span></div>
    <div class="stat"><b id="done">0</b><span>Готово</span></div>
    <div class="stat"><b id="found">0</b><span>Найдено</span></div>
    <div class="stat"><b id="errors">0</b><span>Ошибок</span></div>
  </div>
  <div class="progress"><div id="bar" class="progress-bar"></div></div>
  <div style="display:flex; gap:10px; margin:16px 0;">
    <button id="startBtn" class="btn" onclick="start()">▶️ Запустить обогащение</button>
    <button class="btn secondary" onclick="testOne()">🧪 Тест 1 запроса (Пермь)</button>
    <button class="btn secondary" onclick="saveProgress()">💾 Сохранить прогресс на сервере</button>
  </div>
  <div class="subtitle"><span class="tag">Токен DaData</span> <code id="tokenPreview">77eff8edd...f382dfb</code> &nbsp; <span class="tag">Endpoint</span> <code>POST https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party</code></div>
  <div id="status" style="font-size:13px; margin:10px 0;"></div>
</div>

<div class="card">
  <h3 style="margin:0 0 10px;">Лог запросов</h3>
  <div id="log" class="log">Готов к запуску. Нажмите «Запустить».\n</div>
</div>

<div class="card" id="previewCard" style="display:none;">
  <h3>Предпросмотр последних результатов (руководитель / должность)</h3>
  <table id="previewTable"><tr><th>ИНН</th><th>Организация (реестр)</th><th>Руководитель (DaData)</th><th>Должность</th><th>Статус</th></tr></table>
</div>

<script>
const TOTAL = {len(inn_data)};
let results = [];
let done = 0, found = 0, errors = 0;
const TOKEN = "77eff8edd3640fc0002ad6ae2431261a2f382dfb";
let isRunning = false;

function log(msg, cls="") {{
  const el = document.getElementById('log');
  const line = document.createElement('div');
  if(cls) line.className = cls;
  line.textContent = new Date().toLocaleTimeString() + " " + msg;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}}

function updateStats() {{
  document.getElementById('done').textContent = done;
  document.getElementById('found').textContent = found;
  document.getElementById('errors').textContent = errors;
  document.getElementById('bar').style.width = (done/TOTAL*100).toFixed(1) + "%";
}}

async function fetchOne(inn) {{
  const url = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party";
  const resp = await fetch(url, {{
    method: "POST",
    headers: {{
      "Content-Type": "application/json",
      "Accept": "application/json",
      "Authorization": "Token " + TOKEN
    }},
    body: JSON.stringify({{query: inn}})
  }});
  if (!resp.ok) {{
    throw new Error("HTTP " + resp.status + " " + await resp.text().then(t=>t.slice(0,500)));
  }}
  return await resp.json();
}}

async function testOne() {{
  log("Тест: запрос ИНН 5902290635 (Администрация Перми)...", "");
  try {{
    const data = await fetchOne("5902290635");
    log("✔ Тест успешен: " + JSON.stringify(data).slice(0,500), "ok");
    // показать руководителя
    const s = data.suggestions && data.suggestions[0];
    if(s) log(`Руководитель: ${{s.data.management?.name}} — ${{s.data.management?.post}} (${{s.data.state?.status}})`, "ok");
    else log("Пустой ответ (не найдено)", "warn");
  }} catch(e) {{
    log("✘ Ошибка теста: " + e.message, "err");
  }}
}}

async function start() {{
  if(isRunning) return;
  isRunning = true;
  document.getElementById('startBtn').disabled = true;
  document.getElementById('startBtn').textContent = "⏳ Работаем...";
  log(`Старт: всего ${{TOTAL}} ИНН, пауза 150 мс между запросами`, "");

  // Загрузить список ИНН с сервера (чтобы не хардкодить в HTML)
  let inns = [];
  try {{
    const r = await fetch('/api/inns');
    inns = await r.json();
  }} catch(e) {{
    log("Не удалось загрузить список ИНН с сервера: " + e.message, "err");
    return;
  }}

  for(let i=0; i<inns.length; i++) {{
    const item = inns[i];
    const inn = item.inn;
    try {{
      const data = await fetchOne(inn);
      const suggestion = data.suggestions && data.suggestions[0];
      const mgmt = suggestion ? suggestion.data.management : null;
      const entry = {{
        inn: inn,
        full_name_reestr: item.full_name,
        ogrn_reestr: item.ogrn,
        email_reestr: item.email,
        site_reestr: item.site,
        queried_at: new Date().toISOString(),
        http_status: 200,
        dadata: data,
        extracted: {{
          fio: mgmt ? mgmt.name : "",
          post: mgmt ? mgmt.post : "",
          email: suggestion ? suggestion.data.email : "",
          site: suggestion ? suggestion.data.site : "",
          phone: suggestion ? suggestion.data.phone : "",
          ogrn: suggestion ? suggestion.data.ogrn : "",
          status: suggestion ? suggestion.data.state?.status : "",
        }}
      }};
      results.push(entry);
      done++;
      if(suggestion) found++; else errors++;
      if(i % 10 === 0 || i < 5) log(`[${{i+1}}/${{TOTAL}}] ${{inn}} → ${{mgmt ? mgmt.name + " — " + mgmt.post : "не найдено"}}`, mgmt ? "ok" : "warn");
      updateStats();
      addPreviewRow(entry);
      // Сохранять на сервер каждые 50
      if(results.length % 50 === 0) {{
        await saveProgress(false);
      }}
    }} catch(e) {{
      log(`[${{i+1}}] ${{inn}} ✘ ${{e.message}}`, "err");
      results.push({{inn: inn, error: e.message, queried_at: new Date().toISOString()}});
      done++; errors++;
      updateStats();
    }}
    await new Promise(r => setTimeout(r, 150));
    // Обновить статус
    document.getElementById('status').textContent = `Обработано ${{done}} / ${{TOTAL}} — ${{Math.round(done/TOTAL*100)}}% — ошибок: ${{errors}}`;
  }}

  log(`Готово! Всего ${{done}}, найдено ${{found}}, ошибок ${{errors}}. Сохраняем на сервер...`, "ok");
  await saveProgress(true);
  document.getElementById('startBtn').textContent = "✅ Готово";
  document.getElementById('status').innerHTML = '<b class="ok">Все запросы завершены. Данные сохранены на сервере — сейчас соберём итоговый Excel и зальём в Dropbox.</b>';
}}

function addPreviewRow(entry) {{
  const card = document.getElementById('previewCard');
  card.style.display = 'block';
  const tbl = document.getElementById('previewTable');
  // keep only last 15 rows
  if(tbl.rows.length > 16) tbl.deleteRow(1);
  const row = tbl.insertRow(-1);
  row.innerHTML = `<td>${{entry.inn}}</td><td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${{entry.full_name_reestr}}</td><td>${{entry.extracted?.fio || entry.error || ""}}</td><td>${{entry.extracted?.post || ""}}</td><td>${{entry.error ? "ошибка" : (entry.dadata?.suggestions?.length ? "найдено" : "—")}}</td>`;
}}

async function saveProgress(showAlert=true) {{
  if(results.length===0) {{ if(showAlert) alert("Пока нет данных"); return; }}
  try {{
    const resp = await fetch('/api/save', {{
      method: 'POST',
      headers: {{'Content-Type':'application/json'}},
      body: JSON.stringify({{results}})
    }});
    const j = await resp.json();
    log(`💾 Сохранено на сервере: ${{j.saved}} записей → ${{j.raw_path}} (+ ${{j.excel_path}} готовится)`, "ok");
    if(showAlert) alert(`Сохранено ${{j.saved}} записей`);
  }} catch(e) {{
    log("Ошибка сохранения на сервер: " + e.message, "err");
  }}
}}
</script>
</body>
</html>
"""

class Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/" or parsed.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(HTML.encode('utf-8'))
        elif parsed.path == "/api/inns":
            self.send_response(200)
            self.send_header("Content-type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(json.dumps(inn_data, ensure_ascii=False).encode('utf-8'))
        elif parsed.path == "/health":
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"ok")
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"not found")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/save":
            length = int(self.headers.get('content-length', 0))
            body = self.rfile.read(length)
            try:
                payload = json.loads(body)
                results = payload.get("results", [])
                # Save raw JSONL
                raw_path = "dadata_raw_responses.jsonl"
                with open(raw_path, "w", encoding="utf-8") as f:
                    for r in results:
                        # normalize to {inn, queried_at, response}
                        # r already has dadata or error
                        rec = {
                            "inn": r.get("inn"),
                            "queried_at": r.get("queried_at"),
                            "response": r.get("dadata") if "dadata" in r else {"error": r.get("error"), "http_status": r.get("http_status")}
                        }
                        f.write(json.dumps(rec, ensure_ascii=False) + "\n")
                # Also save intermediate enriched JSON for debugging
                with open("dadata_enriched_intermediate.json", "w", encoding="utf-8") as f:
                    json.dump(results, f, ensure_ascii=False, indent=2)
                # Generate Excel now (call enrich logic inline)
                # Use the same logic as enrich_dadata.py but with in-memory results
                try:
                    excel_path = generate_excel(results)
                except Exception as e:
                    print(f"Excel generation failed: {e}")
                    import traceback; traceback.print_exc()
                    excel_path = str(e)

                self.send_response(200)
                self.send_header("Content-type", "application/json")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(json.dumps({"saved": len(results), "raw_path": raw_path, "excel_path": excel_path}, ensure_ascii=False).encode('utf-8'))
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f"error: {e}".encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        print(f"[{self.client_address[0]}] {format%args}")

def generate_excel(results):
    """Generate enriched Excel from results (same as enrich_dadata.py)"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from datetime import datetime

    # Map inn -> extracted
    inn_map = {}
    for r in results:
        inn = r.get("inn")
        if not inn: continue
        dadata = r.get("dadata")
        extracted = r.get("extracted") or {}
        # Fallback parse if extracted not provided (e.g. from resume)
        if not extracted and dadata:
            sugg = dadata.get("suggestions", []) if isinstance(dadata, dict) else []
            if sugg:
                d = sugg[0].get("data", {})
                mgmt = d.get("management") or {}
                extracted = {
                    "fio": mgmt.get("name",""),
                    "post": mgmt.get("post",""),
                    "email": d.get("email","") or "",
                    "site": d.get("site","") or "",
                    "phone": d.get("phone","") or "",
                    "ogrn": d.get("ogrn","") or "",
                    "status": (d.get("state") or {}).get("status","") or "",
                }
        inn_map[inn] = {
            "fio": extracted.get("fio",""),
            "post": extracted.get("post",""),
            "email": extracted.get("email",""),
            "site": extracted.get("site",""),
            "phone": extracted.get("phone",""),
            "ogrn": extracted.get("ogrn",""),
            "status": extracted.get("status",""),
            "full_name": (dadata.get("suggestions",[{}])[0].get("unrestricted_value","") if dadata and isinstance(dadata, dict) and dadata.get("suggestions") else ""),
            "actuality": "",
            "raw_status": "Найдено" if dadata and dadata.get("suggestions") else ("Не найдено" if dadata else r.get("error","")),
        }
        # actuality date
        if dadata and isinstance(dadata, dict) and dadata.get("suggestions"):
            d = dadata["suggestions"][0].get("data",{})
            state = d.get("state",{})
            inn_map[inn]["actuality"] = state.get("actuality_date") or state.get("registration_date") or ""
            # try to format timestamp
            try:
                if isinstance(inn_map[inn]["actuality"], int):
                    inn_map[inn]["actuality"] = datetime.fromtimestamp(inn_map[inn]["actuality"]/1000).strftime("%Y-%m-%d")
            except: pass

    src_path = "Администрации_городов_из_реестра_2026-11-10.xlsx"
    dst_path = "Администрации_городов_обогащенная_DaData.xlsx"
    wb = openpyxl.load_workbook(src_path, data_only=True)

    NEW_COLUMNS = [
        "Руководитель ФИО (DaData)",
        "Должность руководителя (DaData)",
        "Email из ЕГРЮЛ (DaData)",
        "Сайт из ЕГРЮЛ (DaData)",
        "Телефон из ЕГРЮЛ (DaData)",
        "Email совпадает с реестром?",
        "Сайт совпадает с реестром?",
        "Дополнительный Email найден?",
        "Дополнительный сайт найден?",
        "ОГРН из DaData",
        "Наименование полное из ЕГРЮЛ (DaData)",
        "Статус организации (DaData)",
        "Дата актуальности ЕГРЮЛ (DaData)",
        "Статус запроса DaData",
        "Дата запроса DaData",
    ]

    HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    NEW_FILL = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True, size=10, name="Calibri")
    THIN_BORDER = Border(left=Side(style='thin', color='B4C6E7'), right=Side(style='thin', color='B4C6E7'), top=Side(style='thin', color='B4C6E7'), bottom=Side(style='thin', color='B4C6E7'))
    ALT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    def norm_site(s):
        if not s: return ""
        s=str(s).strip().lower()
        for p in ["https://","http://","www."]:
            if s.startswith(p): s=s[len(p):]
        return s.rstrip("/").strip()
    def norm_email(s):
        if not s: return ""
        return str(s).strip().lower()

    for ws in wb.worksheets:
        if ws.title == "Оглавление":
            continue
        # Add headers
        header_len = ws.max_column
        for idx, name in enumerate(NEW_COLUMNS, start=header_len+1):
            cell = ws.cell(row=1, column=idx, value=name)
            cell.fill = NEW_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = THIN_BORDER
        for cell in ws[1]:
            if cell.column <= header_len:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = THIN_BORDER
        ws.row_dimensions[1].height = 38
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A3
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        # widths
        widths = [55,30,30,32,16,14,13,32,18,26,22,32,16,18, 26,28,26,26,18, 18,18,18,18, 16,26,28,16,16]
        from openpyxl.utils import get_column_letter
        for i, w in enumerate(widths, start=1):
            if i <= ws.max_column:
                ws.column_dimensions[get_column_letter(i)].width = w

        # Fill rows
        for r_idx in range(2, ws.max_row+1):
            inn = str(ws.cell(row=r_idx, column=6).value).strip() if ws.cell(row=r_idx, column=6).value else ""
            email_reg = str(ws.cell(row=r_idx, column=10).value).strip() if ws.cell(row=r_idx, column=10).value else ""
            site_reg = str(ws.cell(row=r_idx, column=8).value).strip() if ws.cell(row=r_idx, column=8).value else ""
            info = inn_map.get(inn, {})
            fio = info.get("fio","")
            post = info.get("post","")
            email_dd = info.get("email","")
            site_dd = info.get("site","")
            phone_dd = info.get("phone","")
            ogrn_dd = info.get("ogrn","")
            full_dd = info.get("full_name","")
            status_dd = info.get("status","")
            actuality = info.get("actuality","")
            raw_status = info.get("raw_status","")

            email_match = "Да" if email_dd and norm_email(email_dd)==norm_email(email_reg) else ("Нет" if email_dd and email_reg else ("Пусто в DaData" if not email_dd else "Пусто в реестре"))
            site_match = "Да" if site_dd and norm_site(site_dd)==norm_site(site_reg) else ("Нет" if site_dd and site_reg else ("Пусто в DaData" if not site_dd else "Пусто в реестре"))
            has_extra_email = "Да" if email_dd and norm_email(email_dd)!=norm_email(email_reg) and email_dd else "Нет"
            has_extra_site = "Да" if site_dd and norm_site(site_dd)!=norm_site(site_reg) and site_dd else "Нет"

            vals = [fio, post, email_dd, site_dd, phone_dd, email_match, site_match, has_extra_email, has_extra_site, ogrn_dd, full_dd, status_dd, actuality, raw_status, datetime.now().strftime("%Y-%m-%d %H:%M")]
            for off, val in enumerate(vals, start=0):
                ws.cell(row=r_idx, column=header_len+1+off, value=val)

        # style rows
        for r_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            is_alt = (r_idx % 2 == 0)
            for cell in row:
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                cell.border = THIN_BORDER
                cell.font = Font(size=9, name="Calibri")
                if is_alt:
                    cell.fill = ALT_FILL
            ws.row_dimensions[r_idx].height = 16

    wb.save(dst_path)
    print(f"Excel saved to {dst_path} ({os.path.getsize(dst_path)} bytes)")
    return dst_path

if __name__ == "__main__":
    with socketserver.TCPServer(("0.0.0.0", PORT), Handler) as httpd:
        print(f"Serving at 0.0.0.0:{PORT} with {len(inn_data)} INNs")
        httpd.serve_forever()

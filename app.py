from __future__ import annotations

import cgi
import html
import json
import os
import shutil
import sys
import threading
import time
import urllib.parse
import uuid
import mimetypes
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
RUNS_DIR = APP_DIR / "runs"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.Lock()


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env_file(APP_DIR / ".env")

from ocr_ledger import TEMPLATE_PATH, process_images, refresh_outputs_from_review_payload, safe_price


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Handwrite Ledger AI</title>
  <style>
    :root { color-scheme: light; font-family: "Microsoft YaHei", Arial, sans-serif; }
    body { margin: 0; background: #f5f7fb; color: #18202b; }
    nav { max-width: 980px; margin: 18px auto 0; padding: 0 20px; box-sizing: border-box; }
    nav a { color: #155e75; font-weight: 600; margin-right: 14px; text-decoration: none; }
    main { max-width: 980px; margin: 32px auto; padding: 0 20px; }
    h1 { font-size: 24px; margin: 0 0 18px; }
    form { background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 22px; }
    .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
    label { display: block; font-size: 13px; color: #435062; margin-bottom: 6px; }
    input, select { box-sizing: border-box; width: 100%; padding: 9px 10px; border: 1px solid #bdc7d5; border-radius: 6px; font-size: 14px; background: white; }
    input[type=file] { padding: 7px; }
    .full { grid-column: 1 / -1; }
    .hint { color: #6b7280; font-size: 12px; margin-top: 6px; line-height: 1.5; }
    button { margin-top: 18px; padding: 10px 18px; border: 0; border-radius: 6px; background: #1f5f8b; color: white; font-size: 15px; cursor: pointer; }
    button:hover { background: #174b70; }
    .card { background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 18px; margin-top: 16px; }
    .error { border-color: #f0b7b7; background: #fff7f7; color: #9f1d1d; }
    .ok a { color: #155e75; font-weight: 600; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { border-bottom: 1px solid #e5e7eb; text-align: left; padding: 8px; vertical-align: top; }
    th { background: #f3f6fa; }
    pre { white-space: pre-wrap; word-break: break-word; background: #f8fafc; padding: 12px; border-radius: 6px; }
    @media (max-width: 720px) { .grid { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<nav><a href="/">新任务</a><a href="/history">任务历史</a></nav>
<main>
  <h1>Handwrite Ledger AI</h1>
  <form action="/run" method="post" enctype="multipart/form-data">
    <div class="grid">
      <div class="full">
        <label>上传图片</label>
        <input name="images" type="file" accept="image/*" multiple required>
        <div class="hint">每张图片会识别 3 次并比对，再由可选视觉 LLM 复核，最终填入模板。</div>
      </div>
      <div>
        <label>农产品品类</label>
        <select name="product_type">
          <option value="小麦">小麦</option>
          <option value="玉米">玉米</option>
        </select>
      </div>
      <div>
        <label>单价</label>
        <input name="unit_price" type="number" step="0.0001" min="0.0001" value="2.42" required>
      </div>
      <div>
        <label>收购日期</label>
        <input name="purchase_date" type="text" placeholder="可空">
      </div>
      <div>
        <label>产品产地</label>
        <input name="origin" type="text" value="本地">
      </div>
      <div>
        <label>是否自产</label>
        <select name="self_sold">
          <option value="是">是</option>
          <option value="否">否</option>
        </select>
      </div>
      <div>
        <label>等级规格</label>
        <input name="grade" type="text" value="1">
      </div>
      <div>
        <label>计量单位</label>
        <input name="unit" type="text" value="公斤">
      </div>
      <div class="full">
        <label>模板文件</label>
        <input name="template" type="file" accept=".xls">
        <div class="hint">不上传时使用默认模板：%%TEMPLATE_PATH%%</div>
      </div>
      <div class="full">
        <label>OCR API Token</label>
        <input name="token" type="password" value="%%TOKEN_VALUE%%" placeholder="也可以通过环境变量 PADDLEOCR_TOKEN 提供">
      </div>
    </div>
    <button type="submit">运行并生成交付物</button>
  </form>
  %%MESSAGE%%
</main>
</body>
</html>
"""


class LedgerHandler(BaseHTTPRequestHandler):
    server_version = "LedgerOCR/1.0"

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            self.send_html(render_index())
            return
        if parsed.path == "/job":
            self.send_html(render_job_page(urllib.parse.parse_qs(parsed.query).get("id", [""])[0]))
            return
        if parsed.path == "/history":
            self.send_html(render_history_page())
            return
        if parsed.path == "/status":
            self.handle_status(parsed.query)
            return
        if parsed.path == "/review":
            self.handle_review(parsed.query)
            return
        if parsed.path == "/asset":
            self.handle_asset(parsed.query)
            return
        if parsed.path == "/download":
            self.handle_download(parsed.query)
            return
        self.send_error(404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/save-review":
            self.handle_save_review(parsed.query)
            return
        if parsed.path != "/run":
            self.send_error(404)
            return
        try:
            job_id = self.handle_run()
            self.send_response(303)
            self.send_header("Location", f"/job?id={urllib.parse.quote(job_id)}")
            self.end_headers()
        except Exception as exc:
            self.send_html(render_index(error_message(exc)), status=500)

    def handle_run(self) -> str:
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
        unit_price = safe_price(get_field(form, "unit_price", ""))
        product_type = get_field(form, "product_type", "小麦")
        if product_type not in {"小麦", "玉米"}:
            raise ValueError("农产品品类只能选择小麦或玉米")

        token = get_field(form, "token", "").strip() or os.environ.get("PADDLEOCR_TOKEN", "").strip()
        if not token:
            raise ValueError("缺少百度飞桨 Token，请在页面填写或设置 PADDLEOCR_TOKEN")

        run_dir = RUNS_DIR / time.strftime("%Y%m%d_%H%M%S")
        upload_dir = run_dir / "uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)

        template_path = save_template(form, run_dir)
        image_paths = save_images(form, upload_dir)
        if not image_paths:
            raise ValueError("没有收到可处理的图片")

        defaults = {
            "origin": get_field(form, "origin", "本地") or "本地",
            "self_sold": get_field(form, "self_sold", "是") or "是",
            "grade": get_field(form, "grade", "1") or "1",
            "unit": get_field(form, "unit", "公斤") or "公斤",
        }
        job_id = uuid.uuid4().hex[:12]
        job = {
            "id": job_id,
            "state": "running",
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "log": [f"已接收 {len(image_paths)} 张图片，开始后台处理"],
            "result": None,
            "error": None,
            "run_dir": str(run_dir),
        }
        with JOBS_LOCK:
            JOBS[job_id] = job
        worker = threading.Thread(
            target=run_job,
            args=(job_id, image_paths, token, template_path, run_dir, unit_price, product_type, defaults, get_field(form, "purchase_date", "")),
            daemon=True,
        )
        worker.start()
        return job_id

    def handle_status(self, query: str) -> None:
        params = urllib.parse.parse_qs(query)
        job_id = params.get("id", [""])[0]
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            payload = dict(job) if job else {"state": "missing", "error": "任务不存在"}
            if payload.get("result"):
                payload["links"] = result_links(payload["result"])
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def handle_download(self, query: str) -> None:
        params = urllib.parse.parse_qs(query)
        raw = params.get("file", [""])[0]
        requested = Path(raw).resolve()
        runs_root = RUNS_DIR.resolve()
        if not str(requested).startswith(str(runs_root)) or not requested.exists():
            self.send_error(404)
            return
        payload = requested.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        filename = requested.name.encode("utf-8")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote_from_bytes(filename)}")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def handle_asset(self, query: str) -> None:
        params = urllib.parse.parse_qs(query)
        raw = params.get("file", [""])[0]
        requested = Path(raw).resolve()
        runs_root = RUNS_DIR.resolve()
        if not str(requested).startswith(str(runs_root)) or not requested.exists():
            self.send_error(404)
            return
        payload = requested.read_bytes()
        content_type = mimetypes.guess_type(str(requested))[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def handle_review(self, query: str) -> None:
        params = urllib.parse.parse_qs(query)
        run_dir = resolve_run_dir(params.get("dir", [""])[0])
        if run_dir is None:
            self.send_error(404)
            return
        review_path = find_review_json(run_dir)
        if review_path is None:
            self.send_error(404)
            return
        payload = json.loads(review_path.read_text(encoding="utf-8"))
        self.send_html(render_review_page(payload, run_dir))

    def handle_save_review(self, query: str) -> None:
        params = urllib.parse.parse_qs(query)
        run_dir = resolve_run_dir(params.get("dir", [""])[0])
        if run_dir is None:
            self.send_json({"ok": False, "error": "任务目录不存在"}, status=404)
            return
        review_path = find_review_json(run_dir)
        if review_path is None:
            self.send_json({"ok": False, "error": "审核明细不存在"}, status=404)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length).decode("utf-8")
            update = json.loads(body or "{}")
            payload = json.loads(review_path.read_text(encoding="utf-8"))
            payload["rows"] = merge_review_rows(payload.get("rows", []), update.get("rows", []))
            payload = refresh_outputs_from_review_payload(payload)
            review_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            self.send_json({
                "ok": True,
                "saved_at": payload.get("saved_at"),
                "links": result_links(payload),
                "rows": review_rows_for_payload(payload),
            })
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, status=500)

    def send_html(self, body: str, status: int = 200) -> None:
        encoded = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def send_json(self, payload: dict, status: int = 200) -> None:
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)


def get_field(form: cgi.FieldStorage, name: str, default: str = "") -> str:
    item = form[name] if name in form else None
    if item is None or isinstance(item, list):
        return default
    value = item.value
    return value if isinstance(value, str) else default


def save_template(form: cgi.FieldStorage, run_dir: Path) -> Path:
    item = form["template"] if "template" in form else None
    if item is None or isinstance(item, list) or not getattr(item, "filename", ""):
        if not TEMPLATE_PATH.exists():
            raise FileNotFoundError(f"默认模板不存在，请上传 .xls 模板或设置 LEDGER_TEMPLATE_PATH：{TEMPLATE_PATH}")
        return TEMPLATE_PATH
    target = run_dir / "template.xls"
    with target.open("wb") as file_handle:
        shutil.copyfileobj(item.file, file_handle)
    return target


def save_images(form: cgi.FieldStorage, upload_dir: Path) -> list[Path]:
    if "images" not in form:
        return []
    items = form["images"]
    if not isinstance(items, list):
        items = [items]
    paths: list[Path] = []
    for index, item in enumerate(items, start=1):
        filename = Path(getattr(item, "filename", "") or f"image_{index}.jpg").name
        suffix = Path(filename).suffix.lower()
        if suffix not in IMAGE_EXTENSIONS:
            continue
        target = upload_dir / f"{index:02d}_{filename}"
        with target.open("wb") as file_handle:
            shutil.copyfileobj(item.file, file_handle)
        paths.append(target)
    return paths


def resolve_run_dir(raw_dir: str) -> Path | None:
    if not raw_dir:
        return None
    run_dir = Path(raw_dir).resolve()
    runs_root = RUNS_DIR.resolve()
    if not str(run_dir).startswith(str(runs_root)) or not run_dir.exists() or not run_dir.is_dir():
        return None
    return run_dir


def find_review_json(run_dir: Path) -> Path | None:
    preferred = run_dir / "审核明细.json"
    if preferred.exists():
        return preferred
    json_files = sorted(run_dir.glob("*.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    return json_files[0] if json_files else None


def merge_review_rows(existing_rows: list[dict], updated_rows: list[dict]) -> list[dict]:
    editable = {"name", "id_no", "address", "phone", "amount_text"}
    by_index = {int(row.get("index", idx + 1)): row for idx, row in enumerate(updated_rows)}
    merged: list[dict] = []
    for idx, row in enumerate(existing_rows, start=1):
        next_row = dict(row)
        update = by_index.get(idx)
        if update:
            for key in editable:
                next_row[key] = str(update.get(key, next_row.get(key, ""))).strip()
        merged.append(next_row)
    return merged


def run_job(
    job_id: str,
    image_paths: list[Path],
    token: str,
    template_path: Path,
    run_dir: Path,
    unit_price: float,
    product_type: str,
    defaults: dict[str, str],
    purchase_date: str,
) -> None:
    with JOBS_LOCK:
        log = JOBS[job_id]["log"]
    try:
        result = process_images(
            image_paths=image_paths,
            token=token,
            template_path=template_path,
            output_dir=run_dir,
            unit_price=unit_price,
            product_type=product_type,
            defaults=defaults,
            purchase_date=purchase_date,
            log=log,
        )
        with JOBS_LOCK:
            JOBS[job_id]["state"] = "done"
            JOBS[job_id]["result"] = result
            JOBS[job_id]["log"].append(f"完成：生成 {result.get('row_count', 0)} 行")
    except Exception as exc:
        with JOBS_LOCK:
            JOBS[job_id]["state"] = "failed"
            JOBS[job_id]["error"] = str(exc)
            JOBS[job_id]["log"].append(f"失败：{exc}")


def render_index(message: str = "") -> str:
    token_value = os.environ.get("PADDLEOCR_TOKEN", "")
    return (
        INDEX_HTML.replace("%%TEMPLATE_PATH%%", html.escape(str(TEMPLATE_PATH)))
        .replace("%%TOKEN_VALUE%%", html.escape(token_value))
        .replace("%%MESSAGE%%", message)
    )


def history_entries() -> list[dict]:
    entries: list[dict] = []
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    for run_dir in sorted((path for path in RUNS_DIR.iterdir() if path.is_dir()), key=lambda path: path.stat().st_mtime, reverse=True):
        review_path = find_review_json(run_dir)
        payload: dict = {}
        if review_path:
            try:
                payload = json.loads(review_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                payload = {}
        settings = payload.get("settings") or {}
        output_raw = payload.get("output") or ""
        review_workbook_raw = payload.get("review_workbook") or ""
        output = Path(output_raw) if output_raw else None
        review_workbook = Path(review_workbook_raw) if review_workbook_raw else None
        entries.append(
            {
                "run_dir": run_dir,
                "name": run_dir.name,
                "mtime": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(run_dir.stat().st_mtime)),
                "product_type": settings.get("product_type", ""),
                "unit_price": settings.get("unit_price", ""),
                "row_count": payload.get("row_count", ""),
                "output": output if output and output.exists() else None,
                "review_workbook": review_workbook if review_workbook and review_workbook.exists() else None,
                "has_review": bool(review_path),
            }
        )
    return entries


def render_history_page() -> str:
    rows = []
    for entry in history_entries():
        run_dir = urllib.parse.quote(str(entry["run_dir"]))
        review_link = f"/review?dir={run_dir}" if entry["has_review"] else ""
        output_link = f"/download?file={urllib.parse.quote(str(entry['output']))}" if entry["output"] else ""
        review_workbook_link = f"/download?file={urllib.parse.quote(str(entry['review_workbook']))}" if entry["review_workbook"] else ""
        links = []
        if review_link:
            links.append(f'<a href="{review_link}">审核</a>')
        if output_link:
            links.append(f'<a href="{output_link}">下载台账</a>')
        if review_workbook_link:
            links.append(f'<a href="{review_workbook_link}">审核表</a>')
        rows.append(
            "<tr>"
            f"<td>{html.escape(entry['name'])}</td>"
            f"<td>{html.escape(entry['mtime'])}</td>"
            f"<td>{html.escape(str(entry['product_type']))}</td>"
            f"<td>{html.escape(str(entry['unit_price']))}</td>"
            f"<td>{html.escape(str(entry['row_count']))}</td>"
            f"<td>{'　'.join(links) if links else '处理中或无结果'}</td>"
            "</tr>"
        )
    body = "\n".join(rows) or '<tr><td colspan="6">暂无任务历史</td></tr>'
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>任务历史</title>
  <style>
    body {{ margin: 0; background: #f5f7fb; color: #18202b; font-family: "Microsoft YaHei", Arial, sans-serif; }}
    nav, main {{ max-width: 1080px; margin: 18px auto; padding: 0 20px; box-sizing: border-box; }}
    nav a {{ color: #155e75; font-weight: 600; margin-right: 14px; text-decoration: none; }}
    .card {{ background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 18px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    th, td {{ border-bottom: 1px solid #e5e7eb; padding: 10px; text-align: left; }}
    th {{ background: #f3f6fa; color: #334155; }}
    a {{ color: #155e75; font-weight: 600; }}
  </style>
</head>
<body>
<nav><a href="/">新任务</a><a href="/history">任务历史</a></nav>
<main>
  <h1>任务历史</h1>
  <section class="card">
    <table>
      <thead><tr><th>任务目录</th><th>更新时间</th><th>品类</th><th>单价</th><th>行数</th><th>操作</th></tr></thead>
      <tbody>{body}</tbody>
    </table>
  </section>
</main>
</body>
</html>"""


def render_job_page(job_id: str) -> str:
    escaped_job = html.escape(job_id)
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>处理进度</title>
  <style>
    body {{ margin: 0; background: #f5f7fb; color: #18202b; font-family: "Microsoft YaHei", Arial, sans-serif; }}
    nav, main {{ max-width: 980px; margin: 18px auto; padding: 0 20px; box-sizing: border-box; }}
    nav a {{ color: #155e75; font-weight: 600; margin-right: 14px; text-decoration: none; }}
    .card {{ background: #fff; border: 1px solid #d9e0ea; border-radius: 8px; padding: 18px; margin-top: 16px; }}
    .bar {{ height: 10px; border-radius: 999px; background: #e5e7eb; overflow: hidden; }}
    .bar span {{ display: block; height: 100%; width: 20%; background: #1f5f8b; animation: pulse 1.4s ease-in-out infinite; }}
    @keyframes pulse {{ 0% {{ margin-left: -20%; }} 100% {{ margin-left: 100%; }} }}
    pre {{ white-space: pre-wrap; word-break: break-word; background: #f8fafc; padding: 12px; border-radius: 6px; min-height: 240px; }}
    a {{ color: #155e75; font-weight: 600; }}
    .failed {{ color: #9f1d1d; }}
  </style>
</head>
<body>
<nav><a href="/">新任务</a><a href="/history">任务历史</a></nav>
<main>
  <h1>处理进度</h1>
  <section class="card">
    <p id="state">任务：{escaped_job}</p>
    <div class="bar" id="bar"><span></span></div>
    <div id="links"></div>
    <pre id="log">正在读取状态...</pre>
  </section>
</main>
<script>
const jobId = {json.dumps(job_id)};
async function refresh() {{
  const res = await fetch(`/status?id=${{encodeURIComponent(jobId)}}`, {{cache: 'no-store'}});
  const data = await res.json();
  const state = document.getElementById('state');
  const log = document.getElementById('log');
  const links = document.getElementById('links');
  state.textContent = `状态：${{data.state || 'unknown'}}`;
  log.textContent = (data.log || []).join('\\n');
  if (data.state === 'done') {{
    document.getElementById('bar').style.display = 'none';
    links.innerHTML = data.links || '';
    return;
  }}
  if (data.state === 'failed' || data.state === 'missing') {{
    document.getElementById('bar').style.display = 'none';
    state.className = 'failed';
    links.textContent = data.error || '处理失败';
    return;
  }}
  setTimeout(refresh, 2000);
}}
refresh();
</script>
</body>
</html>"""


def review_rows_for_payload(payload: dict) -> list[dict]:
    rows = []
    for index, row in enumerate(payload.get("rows", []), start=1):
        crop = Path(row.get("crop_image", ""))
        crop_url = f"/asset?file={urllib.parse.quote(str(crop))}" if crop.exists() else ""
        rows.append(
            {
                "index": index,
                "source_file": row.get("source_file", ""),
                "source_row": row.get("source_row", ""),
                "crop_url": crop_url,
                "name": row.get("name", ""),
                "id_no": row.get("id_no", ""),
                "address": row.get("address", ""),
                "phone": row.get("phone", ""),
                "amount_text": row.get("amount_text", ""),
                "amount": row.get("amount", ""),
                "quantity": row.get("quantity", ""),
                "raw_text": row.get("raw_text", ""),
                "anomalies": row.get("anomalies", []),
                "ocr_note": row.get("ocr_note", ""),
            }
        )
    return rows


def render_review_page(payload: dict, run_dir: Path) -> str:
    rows = review_rows_for_payload(payload)
    rows_json = json.dumps(rows, ensure_ascii=False)
    settings = payload.get("settings") or {}
    try:
        unit_price = float(settings.get("unit_price") or 2.42)
    except (TypeError, ValueError):
        unit_price = 2.42
    save_url = f"/save-review?dir={urllib.parse.quote(str(run_dir))}"
    output = Path(payload.get("output", ""))
    review_workbook = Path(payload.get("review_workbook", ""))
    output_link = f"/download?file={urllib.parse.quote(str(output))}" if output.exists() else ""
    review_workbook_link = f"/download?file={urllib.parse.quote(str(review_workbook))}" if review_workbook.exists() else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>逐行审核</title>
  <style>
    body {{ margin: 0; background: #f5f7fb; color: #172033; font-family: "Microsoft YaHei", Arial, sans-serif; }}
    header {{ min-height: 52px; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 0 18px; background: #17324d; color: #fff; }}
    header a {{ color: #dff4ff; text-decoration: none; margin-left: 12px; }}
    main {{ display: grid; grid-template-columns: 360px 1fr; gap: 14px; padding: 14px; height: calc(100vh - 80px); box-sizing: border-box; }}
    .panel {{ background: #fff; border: 1px solid #d8e0ea; border-radius: 8px; overflow: hidden; }}
    .list {{ overflow: auto; }}
    .tools {{ padding: 10px; border-bottom: 1px solid #e5e7eb; display: flex; gap: 8px; align-items: center; }}
    .tools input {{ flex: 1; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; }}
    .rowbtn {{ width: 100%; text-align: left; border: 0; border-bottom: 1px solid #eef2f7; background: #fff; padding: 10px 12px; cursor: pointer; font-size: 13px; }}
    .rowbtn:hover, .rowbtn.active {{ background: #edf6fb; }}
    .rowbtn.flagged {{ border-left: 5px solid #f5b301; }}
    .rowbtn .meta {{ color: #64748b; font-size: 12px; margin-top: 4px; }}
    .detail {{ display: grid; grid-template-rows: auto 1fr; min-width: 0; }}
    .imagebox {{ padding: 12px; border-bottom: 1px solid #e5e7eb; background: #f8fafc; }}
    .imagebox img {{ width: 100%; max-height: 42vh; object-fit: contain; background: #fff; border: 1px solid #d8e0ea; border-radius: 6px; }}
    .fields {{ padding: 14px; overflow: auto; }}
    .actions {{ display: flex; align-items: center; gap: 10px; margin-bottom: 12px; }}
    .actions button {{ margin: 0; padding: 8px 14px; border: 0; border-radius: 6px; background: #1f5f8b; color: #fff; cursor: pointer; }}
    .actions button:disabled {{ background: #93a4b3; cursor: default; }}
    .save-state {{ color: #64748b; font-size: 13px; }}
    .edit-grid {{ display: grid; grid-template-columns: 1fr 1.55fr 1.55fr 1.05fr .85fr; gap: 10px; align-items: start; }}
    .meta-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 10px; margin-top: 10px; }}
    .field {{ border: 1px solid #d8e0ea; border-radius: 6px; padding: 9px; min-height: 58px; background: #fff; }}
    .field.warn {{ background: #fff8dd; border-color: #f3ca66; }}
    .label {{ color: #64748b; font-size: 12px; margin-bottom: 5px; }}
    .value {{ font-size: 16px; word-break: break-all; }}
    .field input {{ box-sizing: border-box; width: 100%; border: 0; outline: 0; padding: 0; background: transparent; font-size: 16px; color: #172033; }}
    .notes {{ margin-top: 12px; border: 1px solid #d8e0ea; border-radius: 6px; padding: 10px; background: #fff; white-space: pre-wrap; word-break: break-word; }}
    kbd {{ background: #e2e8f0; border-radius: 4px; padding: 2px 5px; }}
    @media (max-width: 1200px) {{ .edit-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
    @media (max-width: 900px) {{ main {{ grid-template-columns: 1fr; height: auto; }} .list {{ max-height: 320px; }} .edit-grid, .meta-grid {{ grid-template-columns: 1fr; }} }}
  </style>
</head>
<body>
<header>
  <strong>逐行审核</strong>
  <span>共 {len(rows)} 行　快捷键：<kbd>↑</kbd>/<kbd>↓</kbd> 切换
    <a href="/">新任务</a><a href="/history">任务历史</a>
    {f'<a href="{output_link}">下载台账</a>' if output_link else ''}
    {f'<a href="{review_workbook_link}">下载审核表</a>' if review_workbook_link else ''}
  </span>
</header>
<main>
  <section class="panel list">
    <div class="tools">
      <input id="search" placeholder="搜索姓名 / 身份证 / 金额">
      <label><input type="checkbox" id="flagOnly"> 只看异常</label>
    </div>
    <div id="rows"></div>
  </section>
  <section class="panel detail">
    <div class="imagebox" id="imagebox"></div>
    <div class="fields">
      <div class="actions">
        <button id="saveBtn" type="button" disabled>保存修改并重导出</button>
        <span class="save-state" id="saveState">未修改</span>
      </div>
      <div id="fields"></div>
      <div class="notes" id="notes"></div>
    </div>
  </section>
</main>
<script>
const allRows = {rows_json};
const saveUrl = {json.dumps(save_url)};
const unitPrice = {json.dumps(unit_price)};
let filtered = allRows.slice();
let current = 0;
let dirty = false;
const rowsEl = document.getElementById('rows');
const fieldsEl = document.getElementById('fields');
const notesEl = document.getElementById('notes');
const imageBox = document.getElementById('imagebox');
const saveBtn = document.getElementById('saveBtn');
const saveState = document.getElementById('saveState');
function rowText(row) {{
  return `${{row.name || ''}} ${{row.id_no || ''}} ${{row.amount_text || ''}} ${{row.address || ''}} ${{row.phone || ''}}`;
}}
function escapeHtml(value) {{
  return String(value ?? '').replace(/[&<>"']/g, ch => ({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}}[ch]));
}}
function parseAmount(value) {{
  const text = String(value ?? '')
    .replace(/[，,]/g, '')
    .replace(/[＋]/g, '+')
    .replace(/元/g, '')
    .trim();
  if (!text) return null;
  if (text.includes('+')) {{
    let total = 0;
    let found = false;
    text.split('+').forEach(part => {{
      const match = part.match(/\\d+(?:\\.\\d+)?/);
      if (match) {{
        total += Number(match[0]);
        found = true;
      }}
    }});
    return found ? total : null;
  }}
  const match = text.match(/\\d+(?:\\.\\d+)?/);
  return match ? Number(match[0]) : null;
}}
function quantityFromAmount(value) {{
  const amount = parseAmount(value);
  if (amount === null || !Number.isFinite(amount) || !Number.isFinite(unitPrice) || unitPrice <= 0) return '';
  return Math.round((amount / unitPrice) * 100) / 100;
}}
function setDirty(value) {{
  dirty = value;
  saveBtn.disabled = !dirty;
  saveState.textContent = dirty ? '有未保存修改' : '已保存';
}}
function applyFilter() {{
  const q = document.getElementById('search').value.trim();
  const flagOnly = document.getElementById('flagOnly').checked;
  filtered = allRows.filter(row => (!flagOnly || (row.anomalies || []).length) && (!q || rowText(row).includes(q)));
  current = Math.min(current, Math.max(0, filtered.length - 1));
  renderList();
  renderDetail();
}}
function renderList() {{
  rowsEl.innerHTML = '';
  filtered.forEach((row, idx) => {{
    const btn = document.createElement('button');
    btn.className = 'rowbtn' + ((row.anomalies || []).length ? ' flagged' : '') + (idx === current ? ' active' : '');
    btn.innerHTML = `<strong>${{row.index}}. ${{row.name || '(空)'}}　${{row.amount_text || ''}}</strong><div class="meta">${{row.source_file}} / 第 ${{row.source_row}} 行</div>`;
    btn.onclick = () => {{ current = idx; renderList(); renderDetail(); }};
    rowsEl.appendChild(btn);
  }});
}}
function renderDetail() {{
  const row = filtered[current];
  if (!row) {{
    imageBox.innerHTML = '<p>没有可显示的记录</p>';
    fieldsEl.innerHTML = '';
    notesEl.textContent = '';
    return;
  }}
  imageBox.innerHTML = row.crop_url ? `<img src="${{row.crop_url}}" alt="row crop">` : '<p>该记录没有裁剪图；请用新版本重新跑一次。</p>';
  fieldsEl.innerHTML = `
    <div class="edit-grid">
      ${{editField(row, 'name', '姓名')}}
      ${{editField(row, 'id_no', '身份证号')}}
      ${{editField(row, 'address', '住址')}}
      ${{editField(row, 'phone', '电话')}}
      ${{editField(row, 'amount_text', '金额')}}
    </div>
    <div class="meta-grid">
      ${{readField('收购数量', row.quantity ?? '', 'quantityValue')}}
      ${{readField('图片', row.source_file || '')}}
      ${{readField('原图行', row.source_row || '')}}
      ${{readField('OCR比对', row.ocr_note || '')}}
    </div>`;
  document.querySelectorAll('[data-field]').forEach(input => {{
    input.addEventListener('input', event => {{
      const field = event.target.dataset.field;
      row[field] = event.target.value;
      if (field === 'amount_text') {{
        row.quantity = quantityFromAmount(row.amount_text);
        const quantityNode = document.getElementById('quantityValue');
        if (quantityNode) {{
          quantityNode.textContent = row.quantity === '' ? '' : String(row.quantity);
        }}
      }}
      setDirty(true);
      renderList();
    }});
  }});
  notesEl.textContent = [
    '异常提示：' + ((row.anomalies || []).join('；') || '无'),
    '原始行文本：' + (row.raw_text || '')
  ].join('\\n\\n');
}}
function editField(row, key, label) {{
  const warning = (row.anomalies || []).join(' ').includes(label.replace('号','').replace('住址','地址').replace('金额','金额'));
  return `<label class="field ${{warning ? 'warn' : ''}}"><div class="label">${{label}}</div><input data-field="${{key}}" value="${{escapeHtml(row[key] ?? '')}}"></label>`;
}}
function readField(label, value, valueId = '') {{
  const idAttr = valueId ? ` id="${{escapeHtml(valueId)}}"` : '';
  return `<div class="field"><div class="label">${{label}}</div><div class="value"${{idAttr}}>${{escapeHtml(value)}}</div></div>`;
}}
async function saveChanges() {{
  saveBtn.disabled = true;
  saveState.textContent = '正在保存...';
  const rows = allRows.map(row => ({{
    index: row.index,
    name: row.name || '',
    id_no: row.id_no || '',
    address: row.address || '',
    phone: row.phone || '',
    amount_text: row.amount_text || ''
  }}));
  const res = await fetch(saveUrl, {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{rows}})
  }});
  const data = await res.json();
  if (!data.ok) {{
    saveState.textContent = '保存失败：' + (data.error || '未知错误');
    saveBtn.disabled = false;
    return;
  }}
  if (Array.isArray(data.rows)) {{
    allRows.splice(0, allRows.length, ...data.rows);
    applyFilter();
  }}
  setDirty(false);
  saveState.textContent = '已保存并重导出：' + (data.saved_at || '');
}}
document.getElementById('search').addEventListener('input', applyFilter);
document.getElementById('flagOnly').addEventListener('change', applyFilter);
saveBtn.addEventListener('click', saveChanges);
document.addEventListener('keydown', event => {{
  if (event.target && event.target.matches('input')) return;
  if (event.key === 'ArrowDown') {{ current = Math.min(current + 1, filtered.length - 1); renderList(); renderDetail(); }}
  if (event.key === 'ArrowUp') {{ current = Math.max(current - 1, 0); renderList(); renderDetail(); }}
}});
applyFilter();
</script>
</body>
</html>"""


def result_links(result: dict) -> str:
    output = Path(result["output"])
    review_workbook = Path(result.get("review_workbook", ""))
    review = next(output.parent.glob("*.json"), output.parent / "审核明细.json")
    output_link = f"/download?file={urllib.parse.quote(str(output))}"
    review_json_link = f"/download?file={urllib.parse.quote(str(review))}"
    review_page_link = f"/review?dir={urllib.parse.quote(str(output.parent))}"
    parts = [f'<a href="{output_link}">下载台账文件（按模板 xls）</a>']
    parts.append(f'<a href="{review_page_link}" target="_blank">打开逐行审核</a>')
    if review_workbook.exists():
        review_link = f"/download?file={urllib.parse.quote(str(review_workbook))}"
        parts.append(f'<a href="{review_link}">下载审核版 xlsx</a>')
    parts.append(f'<a href="{review_json_link}">下载审核明细 JSON</a>')
    summary = f"已生成 {result.get('row_count', 0)} 行。"
    return "<p>" + summary + "</p><p>" + "　".join(parts) + "</p>"


def success_message(result: dict) -> str:
    output = Path(result["output"])
    review_workbook = Path(result.get("review_workbook", ""))
    review = next(output.parent.glob("*.json"), output.parent / "审核明细.json")
    rows = result.get("rows", [])
    output_link = f"/download?file={urllib.parse.quote(str(output))}"
    review_workbook_link = f"/download?file={urllib.parse.quote(str(review_workbook))}" if review_workbook.exists() else ""
    review_link = f"/download?file={urllib.parse.quote(str(review))}"
    review_page_link = f"/review?dir={urllib.parse.quote(str(output.parent))}"
    lines = result.get("log", [])[-12:]
    return f"""
    <section class="card ok">
      <p>已生成 {len(rows)} 行。</p>
      <p><a href="{output_link}">下载台账文件（按模板 xls）</a>　<a href="{review_page_link}" target="_blank">打开逐行审核</a>　{f'<a href="{review_workbook_link}">下载审核版 xlsx</a>　' if review_workbook_link else ''}<a href="{review_link}">下载审核明细 JSON</a></p>
      <h2>最近处理日志</h2>
      <pre>{html.escape(chr(10).join(lines))}</pre>
    </section>
    """


def error_message(exc: Exception) -> str:
    return f'<section class="card error"><strong>处理失败：</strong>{html.escape(str(exc))}</section>'


def main() -> int:
    port = int(os.environ.get("LEDGER_OCR_PORT", "8765"))
    server = create_server(port)
    print(f"台账 OCR 应用已启动：http://127.0.0.1:{server.server_address[1]}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("已停止")
        return 0
    finally:
        server.server_close()


def create_server(port: int = 8765) -> ThreadingHTTPServer:
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    return ThreadingHTTPServer(("127.0.0.1", port), LedgerHandler)


if __name__ == "__main__":
    raise SystemExit(main())

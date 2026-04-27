import os
import threading
import time
import webbrowser
from http.cookies import SimpleCookie
from typing import Any, Dict, List, Tuple
from urllib.parse import urlencode, unquote

import openpyxl
import requests
from flask import Flask, render_template_string, request

BASE_URL = "https://www.importyeti.com"
HOST = "127.0.0.1"
PORT = 5000
SEARCH_ENDPOINTS = {
    "companies": "/api/search",
    "addresses": "/api/search/addresses",
    "hs-codes": "/api/search/hs-codes",
}

REFERER_PATHS = {
    "companies": "/search",
    "addresses": "/search/addresses",
    "hs-codes": "/search/hs-codes",
}

SEARCH_SCOPE_OPTIONS: List[Tuple[str, str]] = [
    ("companies", "公司 /api/search"),
    ("addresses", "地址 /api/search/addresses"),
    ("hs-codes", "HS Code /api/search/hs-codes"),
]

TYPE_OPTIONS: List[Tuple[str, str]] = [
    ("", "Any"),
    ("company", "Company"),
    ("supplier", "Supplier"),
]

RECENT_SHIPMENT_OPTIONS: List[Tuple[str, str]] = [
    ("", "Any"),
    ("1mo", "Last 1 month"),
    ("3mo", "Last 3 months"),
    ("6mo", "Last 6 months"),
    ("1yr", "Last 1 year"),
    ("2yr", "Last 2 years"),
]

SHIPMENT_TOTAL_OPTIONS: List[Tuple[str, str]] = [
    ("", "Any"),
    ("10", "At least 10"),
    ("50", "At least 50"),
    ("100", "At least 100"),
    ("500", "At least 500"),
    ("1000", "At least 1000"),
    ("5000", "At least 5000"),
    ("custom", "Custom"),
]

HTML_TEMPLATE = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ImportYeti Flask 筛选搜索</title>
  <style>
    :root {
      --bg: #eef3f8;
      --card: #ffffff;
      --text: #14243a;
      --muted: #5b6f86;
      --line: #d6e0ea;
      --primary: #136f63;
      --primary-soft: #e8f6f2;
      --danger: #c3362d;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      color: var(--text);
      background:
        radial-gradient(circle at 10% 15%, #ffffff 0%, transparent 45%),
        radial-gradient(circle at 92% 25%, #ffffff 0%, transparent 40%),
        linear-gradient(145deg, #f3f7fb, #dfe9f4);
      min-height: 100vh;
    }
    .wrap { max-width: 1280px; margin: 0 auto; padding: 22px 16px 32px; }
    .card {
      background: var(--card);
      border: 1px solid var(--line);
      border-radius: 14px;
      box-shadow: 0 8px 28px rgba(17, 39, 63, .06);
      padding: 16px;
      margin-bottom: 14px;
    }
    h1 { margin: 0 0 10px; }
    .sub { margin: 0 0 14px; color: var(--muted); }
    .grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(180px, 1fr));
      gap: 10px;
    }
    .field { display: flex; flex-direction: column; gap: 5px; }
    label { font-size: 13px; color: var(--muted); }
    input, select {
      border: 1px solid var(--line);
      border-radius: 10px;
      min-height: 38px;
      padding: 8px 10px;
      font-size: 14px;
      outline: none;
    }
    input:focus, select:focus {
      border-color: var(--primary);
      box-shadow: 0 0 0 3px rgba(19, 111, 99, .15);
    }
    .actions {
      margin-top: 12px;
      display: flex;
      gap: 12px;
      align-items: center;
      flex-wrap: wrap;
    }
    .btn {
      border: 0;
      border-radius: 10px;
      min-height: 38px;
      padding: 0 16px;
      color: #fff;
      background: var(--primary);
      font-weight: 600;
      cursor: pointer;
    }
    .hint { color: var(--muted); font-size: 12px; margin: 0; }
    .ok { color: #166534; background: #ecfdf3; border: 1px solid #c9efdb; border-radius: 10px; padding: 10px; }
    .err { color: var(--danger); font-weight: 600; margin: 0; }
    .chips { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 8px; }
    .chip { border: 1px solid var(--line); background: var(--primary-soft); border-radius: 999px; padding: 6px 10px; font-size: 12px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { border-bottom: 1px solid var(--line); text-align: left; padding: 8px; vertical-align: top; }
    th { background: #f7fafc; color: var(--muted); font-size: 12px; }
    @media (max-width: 1100px) { .grid { grid-template-columns: repeat(2, minmax(180px, 1fr)); } }
    @media (max-width: 640px) { .grid { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>ImportYeti Flask 筛选搜索</h1>
      <p class="sub">可选接口、类型、最近出货、最小出货、国家、分页，以及 Cookie 认证；结果可选导出为 Excel。</p>

      <form method="get">
        <div class="grid">
          <div class="field">
            <label>关键词 q</label>
            <input name="q" value="{{ form.q }}" required>
          </div>

          <div class="field">
            <label>接口 search_scope</label>
            <select name="search_scope">
              {% for v, l in scope_options %}
              <option value="{{ v }}" {% if form.search_scope == v %}selected{% endif %}>{{ l }}</option>
              {% endfor %}
            </select>
          </div>

          <div class="field">
            <label>类型 type</label>
            <select name="type">
              {% for v, l in type_options %}
              <option value="{{ v }}" {% if form.type == v %}selected{% endif %}>{{ l }}</option>
              {% endfor %}
            </select>
          </div>

          <div class="field">
            <label>最近出货 mostrecentshipment</label>
            <select name="mostrecentshipment">
              {% for v, l in recent_options %}
              <option value="{{ v }}" {% if form.mostrecentshipment == v %}selected{% endif %}>{{ l }}</option>
              {% endfor %}
            </select>
          </div>

          <div class="field">
            <label>最小出货 shipmentstotal</label>
            <select name="shipmentstotal">
              {% for v, l in shipment_options %}
              <option value="{{ v }}" {% if form.shipmentstotal == v %}selected{% endif %}>{{ l }}</option>
              {% endfor %}
            </select>
          </div>

          <div class="field">
            <label>自定义最小出货（custom）</label>
            <input name="custom_shipmentstotal" value="{{ form.custom_shipmentstotal }}" placeholder="例如 350">
          </div>

          <div class="field">
            <label>国家 countryCode</label>
            <input name="countryCode" value="{{ form.countryCode }}" placeholder="例如 US">
          </div>

          <div class="field">
            <label>起始页 page</label>
            <input name="page" value="{{ form.page }}">
          </div>

          <div class="field">
            <label>最多页数 max_pages</label>
            <input name="max_pages" value="{{ form.max_pages }}">
          </div>

          <div class="field">
            <label>importyeti_token（可选）</label>
            <input name="token" value="{{ form.token }}" placeholder="可填 token 或整段 Cookie">
          </div>

          <div class="field">
            <label>cf_clearance（可选）</label>
            <input name="cf_clearance" value="{{ form.cf_clearance }}" placeholder="可填 clearance 或整段 Cookie">
          </div>

          <div class="field">
            <label>完整 Cookie（可选）</label>
            <input name="cookie_header" value="{{ form.cookie_header }}" placeholder="importyeti_token=...; cf_clearance=...">
          </div>

          <div class="field">
            <label>Excel 文件名（可选）</label>
            <input name="excel_name" value="{{ form.excel_name }}" placeholder="例如 hang_zhou.xlsx">
          </div>
        </div>

        <div class="actions">
          <label><input type="checkbox" name="export_excel" value="1" {% if form.export_excel %}checked{% endif %}> 导出 Excel</label>
          <button class="btn" type="submit">开始搜索</button>
        </div>
        <p class="hint">提示：403 一般是 Cookie 失效；可把浏览器请求头中的整段 Cookie 直接粘贴到“完整 Cookie”。</p>
      </form>
    </div>

    {% if error %}
    <div class="card"><p class="err">请求失败：{{ error }}</p></div>
    {% endif %}

    {% if message %}
    <div class="card"><div class="ok">{{ message }}</div></div>
    {% endif %}

    {% if summary %}
    <div class="card">
      <div class="chips">
        <span class="chip">mode: {{ summary.mode }}</span>
        <span class="chip">rows: {{ summary.rows }}</span>
        <span class="chip">totalHits: {{ summary.totalHits }}</span>
        <span class="chip">totalPages: {{ summary.totalPages }}</span>
        <span class="chip">totalShipments: {{ summary.totalShipments }}</span>
      </div>

      {% if summary.mode == "companies" %}
      <table>
        <thead>
          <tr>
            <th>#</th>
            <th>title</th>
            <th>type</th>
            <th>country</th>
            <th>shipments</th>
            <th>mostRecent</th>
            <th>address</th>
            <th>url</th>
          </tr>
        </thead>
        <tbody>
          {% for row in items %}
          <tr>
            <td>{{ loop.index }}</td>
            <td>{{ row.title }}</td>
            <td>{{ row.type }}</td>
            <td>{{ row.countryCode }}</td>
            <td>{{ row.totalShipments }}</td>
            <td>{{ row.mostRecentShipment }}</td>
            <td>{{ row.address }}</td>
            <td>{{ row.url }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
      {% else %}
      <table>
        <thead>
          <tr><th>#</th><th>value</th></tr>
        </thead>
        <tbody>
          {% for row in items %}
          <tr>
            <td>{{ loop.index }}</td>
            <td>{{ row.value }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
      {% endif %}
    </div>
    {% endif %}
  </div>
</body>
</html>
"""

app = Flask(__name__)


def to_pos_int(raw: str, default: int) -> int:
    value = (raw or "").strip()
    if value.isdigit() and int(value) > 0:
        return int(value)
    return default


def extract_cookie_value(raw: str, cookie_name: str) -> str:
    value = (raw or "").strip()
    if not value:
        return ""

    # URL decode first to handle form submission encoding
    value = unquote(value)

    cookie = SimpleCookie()
    try:
        cookie.load(value)
        if cookie_name in cookie:
            return cookie[cookie_name].value
    except Exception:
        pass

    prefix = f"{cookie_name}="
    if value.lower().startswith(prefix.lower()):
        return value.split("=", 1)[1].split(";", 1)[0].strip()

    return value


def normalize_company(item: Dict[str, Any]) -> Dict[str, str]:
    return {
        "title": str(item.get("title", "")),
        "countryCode": str(item.get("countryCode", "")),
        "type": str(item.get("type", "")),
        "address": str(item.get("address", "")),
        "totalShipments": str(item.get("totalShipments", "")),
        "mostRecentShipment": str(item.get("mostRecentShipment", "")),
        "url": str(item.get("url", "")),
    }


def parse_generic_list(data: Any) -> List[Dict[str, str]]:
    candidates: List[Any] = []
    if isinstance(data, list):
        candidates = data
    elif isinstance(data, dict):
        for key in ["searchResults", "results", "data", "items", "suggestions"]:
            if isinstance(data.get(key), list):
                candidates = data.get(key)
                break

    rows: List[Dict[str, str]] = []
    for item in candidates:
        if isinstance(item, str):
            rows.append({"value": item})
            continue
        if isinstance(item, dict):
            for field in ["address", "hsCode", "title", "name", "code", "label", "value"]:
                if item.get(field) not in (None, ""):
                    rows.append({"value": str(item.get(field))})
                    break
            else:
                rows.append({"value": str(item)})
            continue
        rows.append({"value": str(item)})
    return rows


class ImportYetiClient:
    def __init__(self, token: str, cf_clearance: str, cookie_header: str):
        self.session = requests.Session()
        self.session.headers.update(
            {
                "accept": "application/json, text/plain, */*",
                "accept-language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                "api_auth": "true",
                "origin": BASE_URL,
                "priority": "u=1, i",
                "sec-ch-ua": '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
                "sec-ch-ua-arch": '"x86"',
                "sec-ch-ua-bitness": '"64"',
                "sec-ch-ua-full-version": '"147.0.7727.116"',
                "sec-ch-ua-full-version-list": '"Google Chrome";v="147.0.7727.116", "Not.A/Brand";v="8.0.0.0", "Chromium";v="147.0.7727.116"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-model": '""',
                "sec-ch-ua-platform": '"macOS"' if os.uname().sysname == "Darwin" else '"Windows"',
                "sec-ch-ua-platform-version": '"14.4.1"' if os.uname().sysname == "Darwin" else '"10.0.0"',
                "sec-fetch-dest": "empty",
                "sec-fetch-mode": "cors",
                "sec-fetch-site": "same-origin",
                "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.7727.116 Safari/537.36" if os.uname().sysname == "Darwin" else "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
            }
        )

        # 优先使用完整 cookie_header，否则分别设置 token 和 clearance
        if cookie_header:
            # 清理 cookie_header：URL decode + 移除前导/末尾空白
            cleaned_cookie = unquote(cookie_header.strip())
            self.session.headers["cookie"] = cleaned_cookie
        else:
            cookies_to_set = {}
            if token:
                cookies_to_set["importyeti_token"] = token
            if cf_clearance:
                cookies_to_set["cf_clearance"] = cf_clearance
            cookies_to_set["_importyeti_returning_user"] = "1"
            
            for name, value in cookies_to_set.items():
                self.session.cookies.set(name, value, domain=".importyeti.com")

    def search(
        self,
        q: str,
        search_scope: str,
        api_params: Dict[str, str],
        start_page: int,
        max_pages: int,
        timeout: int = 30,
        max_retries: int = 2,
    ) -> Dict[str, Any]:
        if search_scope not in SEARCH_ENDPOINTS:
            raise ValueError("search_scope 无效")

        endpoint = SEARCH_ENDPOINTS[search_scope]
        referer_path = REFERER_PATHS[search_scope]
        all_items: List[Dict[str, Any]] = []
        total_pages_seen = start_page
        total_hits = None
        total_shipments = None

        for page in range(start_page, start_page + max_pages):
            params = dict(api_params)
            params["page"] = str(page)

            referer_query = urlencode({k: v for k, v in params.items() if k != "page"})
            self.session.headers["referer"] = f"{BASE_URL}{referer_path}?{referer_query}"

            last_error = None
            for attempt in range(max_retries):
                try:
                    response = self.session.get(f"{BASE_URL}{endpoint}", params=params, timeout=timeout)
                    if response.status_code == 403:
                        last_error = "HTTP 403: Cloudflare 风控。请检查 Cookie 是否有效，或更新 cf_clearance。"
                        if attempt < max_retries - 1:
                            time.sleep(1)
                            continue
                        raise RuntimeError(last_error)
                    elif response.status_code != 200:
                        preview = response.text[:220].replace("\n", " ")
                        raise RuntimeError(f"HTTP {response.status_code}: {preview}")

                    try:
                        data = response.json()
                    except Exception as exc:
                        preview = response.text[:300].replace("\n", " ")
                        raise RuntimeError(f"响应不是 JSON: {preview}") from exc
                    
                    break  # 成功则退出重试循环
                except Exception as exc:
                    last_error = exc
                    if attempt < max_retries - 1:
                        time.sleep(1)
                    else:
                        raise RuntimeError(f"请求失败: {exc}") from exc
            else:
                raise RuntimeError(last_error or "请求失败")

            if search_scope == "companies":
                rows = [normalize_company(x) for x in data.get("searchResults", [])]
            else:
                rows = parse_generic_list(data)

            all_items.extend(rows)
            if isinstance(data, dict):
                total_pages_seen = int(data.get("totalPages") or total_pages_seen)
                total_hits = data.get("totalHits")
                total_shipments = data.get("totalShipments")

            if search_scope != "companies":
                break
            if page >= total_pages_seen:
                break

            time.sleep(0.35)

        return {
            "items": all_items,
            "summary": {
                "mode": search_scope,
                "rows": len(all_items),
                "totalPages": total_pages_seen,
                "totalHits": total_hits,
                "totalShipments": total_shipments,
            },
        }


def save_to_excel(items: List[Dict[str, Any]], file_name: str) -> None:
    if not items:
        return

    fieldnames = sorted({k for row in items for k in row.keys()})
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "importyeti"
    ws.append(fieldnames)
    for row in items:
        ws.append([str(row.get(field, "")) for field in fieldnames])
    wb.save(file_name)


def default_form() -> Dict[str, Any]:
    return {
        "q": "hang zhou",
        "search_scope": "companies",
        "type": "",
        "mostrecentshipment": "",
        "shipmentstotal": "",
        "custom_shipmentstotal": "",
        "countryCode": "",
        "page": "1",
        "max_pages": "1",
        "token": "",
        "cf_clearance": "",
        "cookie_header": "",
        "export_excel": False,
        "excel_name": "importyeti_results.xlsx",
    }


@app.route("/", methods=["GET"])
def index() -> str:
    form = default_form()
    items: List[Dict[str, Any]] = []
    summary = None
    error = ""
    message = ""

    if request.args:
        form["q"] = request.args.get("q", "").strip()
        form["search_scope"] = request.args.get("search_scope", "companies").strip()
        form["type"] = request.args.get("type", "").strip()
        form["mostrecentshipment"] = request.args.get("mostrecentshipment", "").strip()
        form["shipmentstotal"] = request.args.get("shipmentstotal", "").strip()
        form["custom_shipmentstotal"] = request.args.get("custom_shipmentstotal", "").strip()
        form["countryCode"] = request.args.get("countryCode", "").strip().upper()
        form["page"] = request.args.get("page", "1").strip() or "1"
        form["max_pages"] = request.args.get("max_pages", "1").strip() or "1"
        form["token"] = request.args.get("token", "").strip()
        form["cf_clearance"] = request.args.get("cf_clearance", "").strip()
        form["cookie_header"] = request.args.get("cookie_header", "").strip()
        form["excel_name"] = request.args.get("excel_name", "importyeti_results.xlsx").strip()
        form["export_excel"] = request.args.get("export_excel") == "1"

        if form["search_scope"] not in SEARCH_ENDPOINTS:
            form["search_scope"] = "companies"

        if not form["q"]:
            error = "关键词 q 不能为空"
        else:
            token = extract_cookie_value(form["token"], "importyeti_token")
            clearance = extract_cookie_value(form["cf_clearance"], "cf_clearance")
            cookie_header = form["cookie_header"]

            if not cookie_header:
                if "importyeti_token=" in form["token"] or "cf_clearance=" in form["token"]:
                    cookie_header = form["token"]
                if "importyeti_token=" in form["cf_clearance"] or "cf_clearance=" in form["cf_clearance"]:
                    cookie_header = form["cf_clearance"]

            if cookie_header:
                parsed_token = extract_cookie_value(cookie_header, "importyeti_token")
                parsed_clearance = extract_cookie_value(cookie_header, "cf_clearance")
                token = token or parsed_token
                clearance = clearance or parsed_clearance

            api_params = {"q": form["q"]}
            if form["type"]:
                api_params["type"] = form["type"]
            if form["mostrecentshipment"]:
                api_params["mostrecentshipment"] = form["mostrecentshipment"]

            if form["shipmentstotal"] == "custom":
                if form["custom_shipmentstotal"].isdigit():
                    api_params["shipmentstotal"] = form["custom_shipmentstotal"]
            elif form["shipmentstotal"]:
                api_params["shipmentstotal"] = form["shipmentstotal"]

            if form["countryCode"]:
                api_params["countryCode"] = form["countryCode"]

            try:
                client = ImportYetiClient(token=token, cf_clearance=clearance, cookie_header=cookie_header)
                result = client.search(
                    q=form["q"],
                    search_scope=form["search_scope"],
                    api_params=api_params,
                    start_page=to_pos_int(form["page"], 1),
                    max_pages=to_pos_int(form["max_pages"], 1),
                )
                items = result["items"]
                summary = result["summary"]

                if form["export_excel"] and items:
                    excel_name = form["excel_name"] or "importyeti_results.xlsx"
                    save_to_excel(items, excel_name)
                    message = f"已导出 Excel: {excel_name}"
            except Exception as exc:
                error = str(exc)

    return render_template_string(
        HTML_TEMPLATE,
        form=form,
        items=items,
        summary=summary,
        error=error,
        message=message,
        scope_options=SEARCH_SCOPE_OPTIONS,
        type_options=TYPE_OPTIONS,
        recent_options=RECENT_SHIPMENT_OPTIONS,
        shipment_options=SHIPMENT_TOTAL_OPTIONS,
    )


def open_browser() -> None:
    webbrowser.open_new(f"http://{HOST}:{PORT}")


if __name__ == "__main__":
    if os.environ.get("IMPORTYETI_NO_BROWSER") != "1":
        threading.Timer(1.0, open_browser).start()
    app.run(host=HOST, port=PORT, debug=False)

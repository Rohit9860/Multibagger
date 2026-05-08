
# """
# Streamlit Dashboard — Screener.in Quarterly Results
# Run: streamlit run app.py

# Two Excel outputs:
#   1. screener_filtered.xlsx  — filtered companies (dynamic quarters, auto-detected from HTML)
#   2. screener_detailed.xlsx  — per-company deep data fetched from each Screener link
# """

# import streamlit as st
# import pandas as pd
# import requests
# from bs4 import BeautifulSoup
# import re, time, random, io, logging
# from openpyxl import load_workbook, Workbook
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter

# logging.basicConfig(level=logging.INFO)
# log = logging.getLogger(__name__)

# st.set_page_config(page_title="Screener Growth Filter", page_icon="📈", layout="wide")

# st.markdown("""
# <style>
#   @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@500&display=swap');
#   html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
#   .block-container { padding-top: 1.5rem; max-width: 1400px; }
#   .kpi-card {
#     background: #0F172A; border-radius: 12px; padding: 20px 24px;
#     color: white; border: 1px solid #1E293B; margin-bottom: 8px;
#   }
#   .kpi-card .val { font-size: 1.9rem; font-weight: 700; margin: 0;
#     font-family: 'DM Mono', monospace; color: #38BDF8; }
#   .kpi-card .lbl { font-size: 0.75rem; text-transform: uppercase;
#     letter-spacing: 0.08em; opacity: 0.55; margin: 4px 0 0; }
#   .criteria-box {
#     background: #F0F9FF; border-left: 4px solid #0EA5E9;
#     padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.92rem;
#   }
#   .warn-box {
#     background: #FFF7ED; border-left: 4px solid #F97316;
#     padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.88rem;
#   }
# </style>
# """, unsafe_allow_html=True)

# # ── Constants ──────────────────────────────────────────────────────────────────
# BASE_URL    = "https://www.screener.in"
# LOGIN_URL   = f"{BASE_URL}/login/"
# RESULTS_URL = f"{BASE_URL}/results/latest/"

# HEADERS = {
#     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
#     "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
#     "Accept-Language": "en-IN,en;q=0.9",
#     "Connection": "keep-alive",
# }


# # ── Auth ───────────────────────────────────────────────────────────────────────
# def login(username: str, password: str) -> requests.Session:
#     session = requests.Session()
#     session.headers.update(HEADERS)
#     resp = session.get(LOGIN_URL, timeout=15)
#     if resp.status_code != 200:
#         raise RuntimeError(f"Cannot reach login page: HTTP {resp.status_code}")
#     soup = BeautifulSoup(resp.text, "lxml")
#     csrf = ""
#     inp = soup.find("input", {"name": "csrfmiddlewaretoken"})
#     if inp:
#         csrf = inp.get("value", "")
#     else:
#         meta = soup.find("meta", {"name": "csrf-token"})
#         csrf = meta["content"] if meta else session.cookies.get("csrftoken", "")
#     if not csrf:
#         csrf = session.cookies.get("csrftoken", "")
#     r = session.post(LOGIN_URL, data={
#         "username": username, "password": password,
#         "csrfmiddlewaretoken": csrf, "next": "/"
#     }, headers={**HEADERS, "Referer": LOGIN_URL, "Origin": BASE_URL},
#        timeout=15, allow_redirects=True)
#     t = r.text.lower()
#     if "logout" in t or "my account" in t:
#         return session
#     if "invalid" in t or "incorrect" in t:
#         raise RuntimeError("Login failed — check your credentials.")
#     if r.url and "login" not in r.url:
#         return session
#     raise RuntimeError("Login status unclear — could not confirm authentication.")


# # ── Fetch with exponential backoff ─────────────────────────────────────────────
# def fetch(session, url: str) -> str:
#     for attempt in range(6):
#         try:
#             resp = session.get(url, timeout=20)
#             if resp.status_code == 429:
#                 wait = int(resp.headers.get("Retry-After", (2**attempt) * 5))
#                 log.warning(f"429 on {url} — waiting {wait}s")
#                 time.sleep(wait)
#                 continue
#             if resp.status_code == 403:
#                 raise RuntimeError("403 Forbidden — session may have expired.")
#             resp.raise_for_status()
#             return resp.text
#         except RuntimeError:
#             raise
#         except Exception as e:
#             wait = (2**attempt) * 3
#             log.warning(f"Fetch error {url} attempt {attempt+1}: {e} — retry in {wait}s")
#             time.sleep(wait)
#     return ""


# def fetch_page(session, page: int) -> str:
#     url = RESULTS_URL if page == 1 else f"{RESULTS_URL}?p={page}"
#     return fetch(session, url)


# # ── Parsing utils ──────────────────────────────────────────────────────────────
# def parse_num(text) -> float | None:
#     if not text or str(text).strip() in ("-", "—", ""):
#         return None
#     cleaned = re.sub(r"[₹,\s%↑↓⇡⇣+]", "", str(text).strip())
#     try:
#         return float(cleaned)
#     except ValueError:
#         return None


# def get_total_pages(html: str) -> int:
#     soup = BeautifulSoup(html, "lxml")
#     pages = set()
#     for a in soup.find_all("a", href=True):
#         m = re.search(r"[?&]p=(\d+)", a["href"])
#         if m:
#             pages.add(int(m.group(1)))
#     return max(pages) if pages else 1


# def detect_quarters_from_html(html: str) -> list:
#     """
#     Auto-detect the 3 quarter columns from the first data-table.
#     HTML order is newest→oldest; we reverse to chronological (oldest first).
#     """
#     soup = BeautifulSoup(html, "lxml")
#     for table in soup.find_all("table", class_="data-table"):
#         thead = table.find("thead")
#         if not thead:
#             continue
#         cells = [th.get_text(strip=True) for th in thead.find_all("th")]
#         quarters = [
#             c for c in cells
#             if re.match(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}", c)
#         ]
#         if len(quarters) >= 3:
#             return list(reversed(quarters[:3]))  # oldest → newest
#     return []


# # ── Results page parser ────────────────────────────────────────────────────────
# def parse_companies_from_html(html: str, quarters: list) -> list:
#     """
#     Each company card structure:
#       <div class="flex-row ...">   <- heading div with company link
#       <div class="bg-base ...">    <- table wrapper
#         <table class="data-table">
#     """
#     if not html or not quarters:
#         return []
#     soup = BeautifulSoup(html, "lxml")
#     companies = []

#     for table in soup.find_all("table", class_="data-table"):
#         thead = table.find("thead")
#         if not thead:
#             continue
#         header_cells = [th.get_text(strip=True) for th in thead.find_all("th")]
#         if "YOY" not in header_cells:
#             continue

#         quarter_cols = {i: h for i, h in enumerate(header_cells) if i >= 2 and h}

#         company_name = company_link = None
#         wrapper = table.find_parent("div", class_="bg-base")
#         if wrapper:
#             hdiv = wrapper.find_previous_sibling("div")
#             if hdiv:
#                 a = hdiv.find("a", href=lambda h: h and "/company/" in h)
#                 if a:
#                     span = a.find("span", class_="hover-link")
#                     company_name = span.get_text(strip=True) if span else a.get_text(strip=True)
#                     href = a["href"]
#                     company_link = BASE_URL + href if href.startswith("/") else href
#         if not company_name:
#             continue

#         tbody = table.find("tbody")
#         if not tbody:
#             continue

#         record = next((c for c in companies if c["Company"] == company_name), None)
#         if record is None:
#             record = {"Company": company_name, "Link": company_link}
#             companies.append(record)

#         for row in tbody.find_all("tr"):
#             cells = row.find_all("td")
#             if not cells:
#                 continue
#             metric = cells[0].get_text(strip=True)
#             if not metric:
#                 continue
#             for col_idx, quarter in quarter_cols.items():
#                 if col_idx < len(cells):
#                     record[f"{metric}|{quarter}"] = cells[col_idx].get_text(strip=True)

#     return companies


# def extract_metric(record: dict, keyword: str, quarters: list) -> dict:
#     result = {}
#     for q in quarters:
#         for k, v in record.items():
#             if keyword.lower() in k.lower() and q.lower() in k.lower():
#                 result[q] = parse_num(v)
#                 break
#     return result


# def is_strictly_increasing(vals: dict, quarters: list) -> bool:
#     v = [vals.get(q) for q in quarters]
#     if any(x is None for x in v):
#         return False
#     return v[0] < v[1] < v[2]


# # ── Deep company detail scraper ────────────────────────────────────────────────
# def scrape_company_detail(session, company_name: str, link: str) -> dict:
#     """
#     Fetch company Screener page and extract:
#     - Key ratios: ROE, ROCE, P/BV, PE, OPM, Interest Cover, Promoter Holding, Market Cap
#     - Quarterly P&L (last 6 quarters)
#     - Annual P&L (last 3 years)
#     """
#     html = fetch(session, link)
#     if not html:
#         return {"Company": company_name, "Link": link, "Error": "Failed to fetch"}

#     soup = BeautifulSoup(html, "lxml")
#     result = {"Company": company_name, "Link": link}

#     # Key ratios
#     ratio_map = {
#         "Market Cap":       ["market cap", "mcap"],
#         "P/E":              ["stock p/e", "p/e"],
#         "P/BV":             ["price to book", "p/b"],
#         "ROE":              ["return on equity", "roe"],
#         "ROCE":             ["roce", "return on capital"],
#         "OPM":              ["opm", "operating profit margin"],
#         "Promoter Holding": ["promoter holding", "promoter"],
#         "Debt to Equity":   ["debt to equity", "d/e"],
#         "Interest Cover":   ["interest coverage", "interest cover"],
#     }
#     for section_id in ["top-ratios", "company-info"]:
#         section = soup.find(id=section_id)
#         if not section:
#             continue
#         for li in section.find_all("li"):
#             label_el = li.find("span", class_="name")
#             value_el = li.find("span", class_="nowrap") or li.find("span", class_="number")
#             if not label_el or not value_el:
#                 continue
#             label_text = label_el.get_text(strip=True).lower()
#             value_text = value_el.get_text(strip=True)
#             for col_name, keywords in ratio_map.items():
#                 if col_name not in result:
#                     if any(kw in label_text for kw in keywords):
#                         result[col_name] = parse_num(value_text)

#     # Quarterly P&L
#     qsec = soup.find(id="quarters")
#     if qsec:
#         qtable = qsec.find("table")
#         if qtable:
#             thead = qtable.find("thead")
#             tbody = qtable.find("tbody")
#             if thead and tbody:
#                 q_headers = [th.get_text(strip=True) for th in thead.find_all("th")]
#                 q_dates = q_headers[1:]
#                 last_n = min(6, len(q_dates))
#                 q_use = q_dates[-last_n:]
#                 for row in tbody.find_all("tr"):
#                     cells = row.find_all("td")
#                     if not cells:
#                         continue
#                     metric = cells[0].get_text(strip=True)
#                     if metric.lower() in ("sales", "net profit", "ebidt", "eps", "opm %"):
#                         offset = len(q_dates) - last_n + 1
#                         for i, qd in enumerate(q_use):
#                             ci = offset + i
#                             if ci < len(cells):
#                                 result[f"{metric} | {qd}"] = parse_num(cells[ci].get_text(strip=True))

#     # Annual P&L
#     asec = soup.find(id="profit-loss")
#     if asec:
#         atable = asec.find("table")
#         if atable:
#             thead = atable.find("thead")
#             tbody = atable.find("tbody")
#             if thead and tbody:
#                 yr_headers = [th.get_text(strip=True) for th in thead.find_all("th")]
#                 yr_dates = yr_headers[1:]
#                 last_n = min(3, len(yr_dates))
#                 yr_use = yr_dates[-last_n:]
#                 for row in tbody.find_all("tr"):
#                     cells = row.find_all("td")
#                     if not cells:
#                         continue
#                     metric = cells[0].get_text(strip=True)
#                     if metric.lower() in ("sales", "net profit", "opm %"):
#                         offset = len(yr_dates) - last_n + 1
#                         for i, yr in enumerate(yr_use):
#                             ci = offset + i
#                             if ci < len(cells):
#                                 result[f"Annual {metric} {yr}"] = parse_num(cells[ci].get_text(strip=True))

#     return result


# # ── Excel builders ─────────────────────────────────────────────────────────────
# def _hdr_style(ws, row=1):
#     hfill = PatternFill("solid", fgColor="1F4E79")
#     hfont = Font(bold=True, color="FFFFFF", name="Arial", size=10)
#     thin  = Side(style="thin", color="B0B0B0")
#     bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
#     for cell in ws[row]:
#         cell.font = hfont; cell.fill = hfill
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border = bdr
#     ws.row_dimensions[row].height = 32


# def _row_style(ws, start=2):
#     gfill = PatternFill("solid", fgColor="E2EFDA")
#     afill = PatternFill("solid", fgColor="F0F7FF")
#     thin  = Side(style="thin", color="D0D0D0")
#     bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
#     for ri, row in enumerate(ws.iter_rows(min_row=start), start=start):
#         fill = gfill if ri % 2 == 0 else afill
#         for cell in row:
#             cell.fill = fill
#             cell.alignment = Alignment(horizontal="left", vertical="center")
#             cell.border = bdr


# def _autowidth(ws, max_w=42):
#     for col_cells in ws.columns:
#         w = max((len(str(c.value or "")) for c in col_cells), default=10)
#         ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(w + 4, max_w)


# def _hyperlink_col(ws, col_name="Link"):
#     for c in ws[1]:
#         if str(c.value) == col_name:
#             col = c.column
#             for ri in range(2, ws.max_row + 1):
#                 cell = ws.cell(row=ri, column=col)
#                 if cell.value and str(cell.value).startswith("http"):
#                     cell.hyperlink = cell.value
#                     cell.font = Font(color="0563C1", underline="single", name="Arial", size=10)
#             return


# def build_filtered_excel(df: pd.DataFrame, quarters: list) -> bytes:
#     buf = io.BytesIO()
#     df.to_excel(buf, index=False, sheet_name="Filtered Companies")
#     buf.seek(0)
#     wb = load_workbook(buf)
#     ws = wb.active
#     _hdr_style(ws)
#     _row_style(ws)
#     _autowidth(ws)
#     _hyperlink_col(ws, "Link")

#     # Color growth columns
#     for col_cells in ws.columns:
#         hdr = str(ws.cell(row=1, column=col_cells[0].column).value or "")
#         if "Growth %" in hdr:
#             for ri in range(2, ws.max_row + 1):
#                 cell = ws.cell(row=ri, column=col_cells[0].column)
#                 try:
#                     v = float(cell.value)
#                     cell.fill = PatternFill("solid", fgColor="C6EFCE" if v >= 20 else ("FFEB9C" if v >= 0 else "FFC7CE"))
#                     cell.font = Font(
#                         color="276221" if v >= 20 else ("9C5700" if v >= 0 else "9C0006"),
#                         bold=v >= 20, name="Arial", size=10
#                     )
#                     cell.number_format = "0.00"
#                 except Exception:
#                     pass

#     ws.freeze_panes = "A2"

#     # Summary sheet
#     ws2 = wb.create_sheet("Summary")
#     ws2["A1"] = "Metric"; ws2["B1"] = "Value"
#     _hdr_style(ws2)
#     rows = [
#         ("Total Qualifying Companies", len(df)),
#         ("Quarters Used",              " → ".join(quarters)),
#         ("Avg Sales Growth %",         round(df["Sales Growth %"].mean(), 2) if "Sales Growth %" in df.columns else "—"),
#         ("Avg Profit Growth %",        round(df["Profit Growth %"].mean(), 2) if "Profit Growth %" in df.columns else "—"),
#         ("Max Sales Growth %",         round(df["Sales Growth %"].max(), 2) if "Sales Growth %" in df.columns else "—"),
#         ("Max Profit Growth %",        round(df["Profit Growth %"].max(), 2) if "Profit Growth %" in df.columns else "—"),
#         ("Top Company (Profit Growth)",df.nlargest(1,"Profit Growth %")["Company"].values[0] if len(df) else "—"),
#     ]
#     for i, (k, v) in enumerate(rows, start=2):
#         ws2[f"A{i}"] = k; ws2[f"B{i}"] = v
#     ws2.column_dimensions["A"].width = 34
#     ws2.column_dimensions["B"].width = 22

#     out = io.BytesIO(); wb.save(out); return out.getvalue()


# def build_detailed_excel(records: list) -> bytes:
#     if not records:
#         wb = Workbook(); ws = wb.active; ws["A1"] = "No data"
#         out = io.BytesIO(); wb.save(out); return out.getvalue()

#     df = pd.DataFrame(records)

#     ratio_cols = ["Company", "Link", "Market Cap", "P/E", "P/BV", "ROE", "ROCE",
#                   "OPM", "Interest Cover", "Debt to Equity", "Promoter Holding", "Error"]
#     ratio_cols = [c for c in ratio_cols if c in df.columns]
#     df_ratios  = df[ratio_cols].copy()

#     qtr_pat = re.compile(r".+\|\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}")
#     qtr_cols = ["Company"] + [c for c in df.columns if qtr_pat.match(str(c))]
#     df_qtr   = df[[c for c in qtr_cols if c in df.columns]].copy()

#     ann_cols = ["Company"] + [c for c in df.columns if str(c).startswith("Annual ")]
#     df_ann   = df[[c for c in ann_cols if c in df.columns]].copy()

#     buf = io.BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as writer:
#         df_ratios.to_excel(writer, index=False, sheet_name="Key Ratios")
#         df_qtr.to_excel(writer, index=False, sheet_name="Quarterly P&L")
#         df_ann.to_excel(writer, index=False, sheet_name="Annual P&L")
#     buf.seek(0)

#     wb = load_workbook(buf)
#     for sname in wb.sheetnames:
#         ws = wb[sname]
#         _hdr_style(ws); _row_style(ws); _autowidth(ws)
#         if sname == "Key Ratios":
#             _hyperlink_col(ws, "Link")
#         # Color ROE / ROCE / OPM
#         thresholds = {"ROE": 15, "ROCE": 15, "OPM": 10}
#         for col_cells in ws.columns:
#             hdr = str(ws.cell(row=1, column=col_cells[0].column).value or "")
#             if hdr in thresholds:
#                 thr = thresholds[hdr]
#                 for ri in range(2, ws.max_row + 1):
#                     cell = ws.cell(row=ri, column=col_cells[0].column)
#                     try:
#                         v = float(cell.value)
#                         good = v >= thr
#                         cell.fill = PatternFill("solid", fgColor="C6EFCE" if good else "FFC7CE")
#                         cell.font = Font(color="276221" if good else "9C0006", bold=True, name="Arial", size=10)
#                         cell.number_format = "0.0"
#                     except Exception:
#                         pass
#         ws.freeze_panes = "A2"

#     out = io.BytesIO(); wb.save(out); return out.getvalue()


# # ── Core orchestrator ──────────────────────────────────────────────────────────
# def run_full_scrape(session, max_pages, delay, progress_bar, status_text):
#     status_text.text("📄 Fetching page 1...")
#     first_html = fetch_page(session, 1)
#     if not first_html:
#         raise RuntimeError("Could not fetch page 1.")

#     quarters = detect_quarters_from_html(first_html)
#     if not quarters:
#         raise RuntimeError("Could not detect quarters from HTML — check debug.html for structure.")
#     status_text.text(f"📅 Quarters detected: {' → '.join(quarters)}")

#     total = get_total_pages(first_html)
#     if max_pages and max_pages > 0:
#         total = min(total, max_pages)

#     all_companies = []
#     progress_bar.progress(1 / max(total, 1), text=f"Page 1 / {total}")
#     all_companies.extend(parse_companies_from_html(first_html, quarters))

#     for page in range(2, total + 1):
#         sleep_time = delay * random.uniform(0.8, 1.5)
#         status_text.text(f"🌐 Page {page}/{total} — {len(all_companies)} companies so far...")
#         time.sleep(sleep_time)
#         html = fetch_page(session, page)
#         all_companies.extend(parse_companies_from_html(html, quarters))
#         progress_bar.progress(page / total, text=f"Page {page} / {total}")

#     status_text.text(f"✅ Scraped {len(all_companies)} companies. Filtering...")

#     qualifying = []
#     for comp in all_companies:
#         sales  = extract_metric(comp, "Sales",      quarters)
#         profit = extract_metric(comp, "Net profit", quarters)
#         if not is_strictly_increasing(sales, quarters):
#             continue
#         if not is_strictly_increasing(profit, quarters):
#             continue
#         s = [sales[q]  for q in quarters]
#         p = [profit[q] for q in quarters]
#         if not s[0] or not p[0]:
#             continue
#         row = {"Company": comp["Company"], "Link": comp["Link"]}
#         for i, q in enumerate(quarters):
#             row[f"Sales {q}"]      = s[i]
#             row[f"Net Profit {q}"] = p[i]
#         row["Sales Growth %"]  = round((s[2] - s[0]) / abs(s[0]) * 100, 2)
#         row["Profit Growth %"] = round((p[2] - p[0]) / abs(p[0]) * 100, 2)
#         qualifying.append(row)

#     return pd.DataFrame(qualifying), quarters


# # ══════════════════════════════════════════════════════════════════════════════
# #   UI
# # ══════════════════════════════════════════════════════════════════════════════

# st.title("📈 Screener Quarterly Growth Filter")
# st.markdown("""
# <div class="criteria-box">
#   <b>Filter:</b> Companies where <b>Sales</b> AND <b>Net Profit</b> are strictly increasing
#   across the 3 most recent quarters — <b>quarters are auto-detected live from Screener.in HTML,
#   never hardcoded</b>. Produces two formatted Excel files.
# </div>
# """, unsafe_allow_html=True)

# with st.expander("🔐 Screener.in Login  (required)", expanded=True):
#     st.markdown("""<div class="warn-box">
#       Login required — authenticated sessions receive full data and avoid 429 rate limits.
#     </div>""", unsafe_allow_html=True)
#     c1, c2 = st.columns(2)
#     with c1:
#         username = st.text_input("Email / Username", placeholder="you@example.com")
#     with c2:
#         password = st.text_input("Password", type="password")

# col_a, col_b, col_c, col_d = st.columns(4)
# with col_a:
#     max_pages = st.number_input("Max pages (0 = all)", min_value=0, max_value=200, value=0)
# with col_b:
#     delay = st.slider("Delay between pages (sec)", 1.0, 6.0, 3.0, 0.5)
# with col_c:
#     fetch_detail = st.checkbox("Fetch deep detail per company", value=True,
#         help="Visits each company Screener page — extracts ROE, ROCE, OPM, quarterly & annual P&L")
# with col_d:
#     detail_delay = st.slider("Detail fetch delay (sec)", 0.5, 4.0, 1.5, 0.5)

# st.markdown("<br>", unsafe_allow_html=True)
# start = st.button("🚀 Start Scraping", type="primary")

# for key in ("df_result", "quarters", "excel1", "excel2"):
#     if key not in st.session_state:
#         st.session_state[key] = None

# if start:
#     if not username or not password:
#         st.error("Please enter your Screener.in credentials.")
#     else:
#         progress_bar = st.progress(0, text="Starting...")
#         status_text  = st.empty()
#         try:
#             status_text.text("🔐 Logging in...")
#             session = login(username, password)
#             status_text.text("✅ Logged in.")

#             df, quarters = run_full_scrape(session, max_pages, delay, progress_bar, status_text)
#             st.session_state.df_result = df
#             st.session_state.quarters  = quarters

#             status_text.text("📊 Building Excel 1 — Filtered Summary...")
#             st.session_state.excel1 = build_filtered_excel(df, quarters)

#             if fetch_detail and not df.empty:
#                 detail_records = []
#                 total_d = len(df)
#                 dp = st.progress(0, text="Fetching company details...")
#                 for i, row in enumerate(df.itertuples(), 1):
#                     dp.progress(i / total_d, text=f"Detail {i}/{total_d}: {row.Company}")
#                     status_text.text(f"🔍 Fetching: {row.Company}")
#                     rec = scrape_company_detail(session, row.Company, row.Link)
#                     detail_records.append(rec)
#                     time.sleep(detail_delay * random.uniform(0.7, 1.3))
#                 dp.progress(1.0, text="Details done!")
#                 status_text.text("📊 Building Excel 2 — Deep Analysis...")
#                 st.session_state.excel2 = build_detailed_excel(detail_records)
#             else:
#                 st.session_state.excel2 = None

#             progress_bar.progress(1.0, text="All done!")
#             status_text.success(f"✅ {len(df)} qualifying companies. Download your Excel files below.")

#         except Exception as e:
#             progress_bar.empty()
#             status_text.error(f"❌ {e}")
#             import traceback
#             st.code(traceback.format_exc())

# # ── Results display ────────────────────────────────────────────────────────────
# df       = st.session_state.df_result
# quarters = st.session_state.quarters or []

# if df is not None:
#     if df.empty:
#         st.warning("No companies matched the filter criteria.")
#     else:
#         st.markdown("---")

#         if quarters:
#             st.info(f"**Quarters (auto-detected):** `{'  →  '.join(quarters)}`")

#         # KPI cards
#         avg_sg = df["Sales Growth %"].mean()  if "Sales Growth %" in df.columns else 0
#         avg_pg = df["Profit Growth %"].mean() if "Profit Growth %" in df.columns else 0
#         top_co = df.nlargest(1, "Profit Growth %")["Company"].values[0] if len(df) else "—"
#         top_pg = df["Profit Growth %"].max()  if "Profit Growth %" in df.columns else 0

#         for col, (val, lbl) in zip(st.columns(4), [
#             (len(df),           "Qualifying Companies"),
#             (f"{avg_sg:.1f}%",  "Avg Sales Growth"),
#             (f"{avg_pg:.1f}%",  "Avg Profit Growth"),
#             (f"{top_pg:.1f}%",  f"Best: {top_co}"),
#         ]):
#             with col:
#                 st.markdown(
#                     f'<div class="kpi-card"><p class="val">{val}</p><p class="lbl">{lbl}</p></div>',
#                     unsafe_allow_html=True,
#                 )

#         st.markdown("<br>", unsafe_allow_html=True)

#         # Downloads
#         dl1, dl2, dl3 = st.columns(3)
#         with dl1:
#             st.download_button("⬇️ Download CSV",
#                 df.to_csv(index=False).encode("utf-8"),
#                 "screener_filtered.csv", "text/csv", use_container_width=True)
#         with dl2:
#             if st.session_state.excel1:
#                 st.download_button("⬇️ Excel 1 — Filtered Summary",
#                     st.session_state.excel1, "screener_filtered.xlsx",
#                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                     use_container_width=True)
#         with dl3:
#             if st.session_state.excel2:
#                 st.download_button("⬇️ Excel 2 — Deep Company Analysis",
#                     st.session_state.excel2, "screener_detailed.xlsx",
#                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                     use_container_width=True)

#         st.markdown("---")

#         # Searchable table
#         search = st.text_input("🔍 Search company", "")
#         fdf = df[df["Company"].str.contains(search, case=False, na=False)] if search else df.copy()
#         st.caption(f"Showing {len(fdf)} of {len(df)} companies")

#         display = fdf.copy()
#         display["Link"] = display["Link"].apply(
#             lambda u: f'<a href="{u}" target="_blank">🔗 View</a>' if pd.notna(u) else "")

#         def color_growth(val):
#             try:
#                 v = float(val)
#                 if v >= 20:  return "background-color:#C6EFCE;color:#276221;font-weight:600"
#                 elif v >= 0: return "background-color:#FFEB9C;color:#9C5700"
#                 return "background-color:#FFC7CE;color:#9C0006"
#             except Exception:
#                 return ""

#         num_fmt = {c: "{:,.1f}" for c in display.columns if any(q in c for q in quarters)}
#         num_fmt.update({c: "{:.2f}%" for c in display.columns if "Growth %" in c})
#         growth_cols = [c for c in display.columns if "Growth %" in c]
#         styled = display.style.applymap(color_growth, subset=growth_cols).format(num_fmt, na_rep="—")
#         st.write(styled.to_html(escape=False), unsafe_allow_html=True)

#         # Charts
#         st.markdown("---")
#         st.markdown("### 📊 Top Companies by Growth")
#         top_n = min(15, len(fdf))
#         ch1, ch2 = st.columns(2)
#         with ch1:
#             st.bar_chart(fdf.nlargest(top_n, "Sales Growth %").set_index("Company")["Sales Growth %"])
#             st.caption(f"Top {top_n} by Sales Growth %")
#         with ch2:
#             st.bar_chart(fdf.nlargest(top_n, "Profit Growth %").set_index("Company")["Profit Growth %"])
#             st.caption(f"Top {top_n} by Profit Growth %")

#         try:
#             import plotly.express as px
#             st.markdown("### 🔵 Sales Growth vs Profit Growth")
#             fig = px.scatter(fdf, x="Sales Growth %", y="Profit Growth %",
#                 hover_name="Company", color="Profit Growth %",
#                 color_continuous_scale="Blues", template="plotly_white")
#             fig.update_traces(marker=dict(size=9, opacity=0.8))
#             fig.update_layout(height=400, margin=dict(l=20, r=20, t=20, b=20))
#             st.plotly_chart(fig, use_container_width=True)
#         except ImportError:
#             pass

# else:
#     st.info("👆 Enter credentials, configure options, and click **Start Scraping**.")
#     with st.expander("ℹ️ How it works & what you get"):
#         st.markdown("""
# **Step 1 — Scrape & Filter**
# - Logs into Screener.in, fetches `/results/latest/` across all pages
# - **Quarters auto-detected from live HTML** — works even when Screener changes the dates
# - Keeps companies where Sales AND Net Profit are both strictly increasing Q1 → Q2 → Q3

# **Excel 1 — Filtered Summary** (`screener_filtered.xlsx`)
# - Dynamic quarter columns (whatever Screener currently shows)
# - Sales & Net Profit per quarter + growth %
# - Color-coded growth cells (green ≥ 20%, yellow ≥ 0%, red < 0%)
# - Summary stats sheet | Hyperlinked Screener links

# **Excel 2 — Deep Company Analysis** (`screener_detailed.xlsx`)
# - Visits each qualifying company's Screener page individually
# - **Sheet 1 — Key Ratios:** ROE, ROCE, P/BV, P/E, OPM, Interest Cover, Debt/Equity, Promoter Holding, Market Cap
# - **Sheet 2 — Quarterly P&L:** Last 6 quarters of Sales, Net Profit, EBIDT, EPS, OPM %
# - **Sheet 3 — Annual P&L:** Last 3 years of Sales, Net Profit, OPM %
# - ROE/ROCE/OPM cells auto-colored green (above threshold) or red

# **Tips**
# - Test with Max pages = 3 first
# - Keep delay ≥ 3s to avoid 429 errors
# - Deep detail adds ~1-2s per company
#         """)

################################################## SECOND ####################################################



















# """
# app.py — Streamlit UI for the Screener Growth Filter.
# All scraping / parsing / Excel logic lives in scraper.py.
# """

# import time
# import random

# import pandas as pd
# import streamlit as st

# from scraper import (
#     login,
#     run_full_scrape,
#     scrape_company_detail,
#     build_filtered_excel,
#     build_detailed_excel,
# )

# # ── Page config ────────────────────────────────────────────────────────────────
# st.set_page_config(page_title="Screener Growth Filter", page_icon="📈", layout="wide")

# st.markdown("""
# <style>
#   @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@500&display=swap');
#   html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
#   .block-container { padding-top: 1.5rem; max-width: 1400px; }

#   .kpi-card {
#     background: #0F172A; border-radius: 12px; padding: 20px 24px;
#     color: white; border: 1px solid #1E293B; margin-bottom: 8px;
#   }
#   .kpi-card .val {
#     font-size: 1.9rem; font-weight: 700; margin: 0;
#     font-family: 'DM Mono', monospace; color: #38BDF8;
#   }
#   .kpi-card .lbl {
#     font-size: 0.75rem; text-transform: uppercase;
#     letter-spacing: 0.08em; opacity: 0.55; margin: 4px 0 0;
#   }
#   .criteria-box {
#     background: #F0F9FF; border-left: 4px solid #0EA5E9;
#     padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.92rem;
#   }
#   .warn-box {
#     background: #FFF7ED; border-left: 4px solid #F97316;
#     padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.88rem;
#   }
# </style>
# """, unsafe_allow_html=True)

# # ── Header ─────────────────────────────────────────────────────────────────────
# st.title("📈 Screener Quarterly Growth Filter")
# st.markdown("""
# <div class="criteria-box">
#   <b>Filter:</b> Companies where <b>Sales</b> AND <b>Net Profit</b> are strictly increasing
#   across the 3 most recent quarters — <b>quarters are auto-detected live from Screener.in HTML,
#   never hardcoded</b>. Produces two formatted Excel files.
# </div>
# """, unsafe_allow_html=True)

# # ── Credentials ────────────────────────────────────────────────────────────────
# with st.expander("🔐 Screener.in Login  (required)", expanded=True):
#     st.markdown("""<div class="warn-box">
#       Login required — authenticated sessions receive full data and avoid 429 rate limits.
#     </div>""", unsafe_allow_html=True)
#     c1, c2 = st.columns(2)
#     with c1:
#         username = st.text_input("Email / Username", placeholder="you@example.com")
#     with c2:
#         password = st.text_input("Password", type="password")

# # ── Options ────────────────────────────────────────────────────────────────────
# col_a, col_b, col_c, col_d = st.columns(4)
# with col_a:
#     max_pages = st.number_input("Max pages (0 = all)", min_value=0, max_value=200, value=0)
# with col_b:
#     delay = st.slider("Delay between pages (sec)", 1.0, 6.0, 3.0, 0.5)
# with col_c:
#     fetch_detail = st.checkbox(
#         "Fetch deep detail per company", value=True,
#         help="Visits each company Screener page — extracts ROE, ROCE, OPM, quarterly & annual P&L",
#     )
# with col_d:
#     detail_delay = st.slider("Detail fetch delay (sec)", 0.5, 4.0, 1.5, 0.5)

# st.markdown("<br>", unsafe_allow_html=True)
# start = st.button("🚀 Start Scraping", type="primary")

# # ── Session-state init ─────────────────────────────────────────────────────────
# for key in ("df_result", "quarters", "excel1", "excel2"):
#     if key not in st.session_state:
#         st.session_state[key] = None

# # ── Scrape trigger ─────────────────────────────────────────────────────────────
# if start:
#     if not username or not password:
#         st.error("Please enter your Screener.in credentials.")
#     else:
#         progress_bar = st.progress(0, text="Starting...")
#         status_text  = st.empty()

#         try:
#             status_text.text("🔐 Logging in...")
#             session = login(username, password)
#             status_text.text("✅ Logged in.")

#             # Progress callbacks wiring Streamlit UI to scraper.py
#             def on_page_start(page, total, n):
#                 status_text.text(f"🌐 Page {page}/{total} — {n} companies so far...")

#             def on_page_done(page, total, n):
#                 progress_bar.progress(page / max(total, 1), text=f"Page {page} / {total}")

#             df, quarters = run_full_scrape(
#                 session, max_pages, delay,
#                 on_page_start=on_page_start,
#                 on_page_done=on_page_done,
#             )
#             st.session_state.df_result = df
#             st.session_state.quarters  = quarters

#             status_text.text("📊 Building Excel 1 — Filtered Summary...")
#             st.session_state.excel1 = build_filtered_excel(df, quarters)

#             if fetch_detail and not df.empty:
#                 detail_records = []
#                 total_d = len(df)
#                 dp = st.progress(0, text="Fetching company details...")
#                 for i, row in enumerate(df.itertuples(), 1):
#                     dp.progress(i / total_d, text=f"Detail {i}/{total_d}: {row.Company}")
#                     status_text.text(f"🔍 Fetching: {row.Company}")
#                     rec = scrape_company_detail(session, row.Company, row.Link)
#                     detail_records.append(rec)
#                     time.sleep(detail_delay * random.uniform(0.7, 1.3))
#                 dp.progress(1.0, text="Details done!")
#                 status_text.text("📊 Building Excel 2 — Deep Analysis...")
#                 st.session_state.excel2 = build_detailed_excel(detail_records)
#             else:
#                 st.session_state.excel2 = None

#             progress_bar.progress(1.0, text="All done!")
#             status_text.success(
#                 f"✅ {len(df)} qualifying companies. Download your Excel files below."
#             )

#         except Exception as e:
#             progress_bar.empty()
#             status_text.error(f"❌ {e}")
#             import traceback
#             st.code(traceback.format_exc())

# # ── Results display ────────────────────────────────────────────────────────────
# df       = st.session_state.df_result
# quarters = st.session_state.quarters or []

# if df is not None:
#     if df.empty:
#         st.warning("No companies matched the filter criteria.")
#     else:
#         st.markdown("---")

#         if quarters:
#             st.info(f"**Quarters (auto-detected):** `{'  →  '.join(quarters)}`")

#         # KPI cards
#         avg_sg = df["Sales Growth %"].mean()  if "Sales Growth %" in df.columns else 0
#         avg_pg = df["Profit Growth %"].mean() if "Profit Growth %" in df.columns else 0
#         top_co = df.nlargest(1, "Profit Growth %")["Company"].values[0] if len(df) else "—"
#         top_pg = df["Profit Growth %"].max()  if "Profit Growth %" in df.columns else 0

#         for col, (val, lbl) in zip(st.columns(4), [
#             (len(df),           "Qualifying Companies"),
#             (f"{avg_sg:.1f}%",  "Avg Sales Growth"),
#             (f"{avg_pg:.1f}%",  "Avg Profit Growth"),
#             (f"{top_pg:.1f}%",  f"Best: {top_co}"),
#         ]):
#             with col:
#                 st.markdown(
#                     f'<div class="kpi-card">'
#                     f'<p class="val">{val}</p>'
#                     f'<p class="lbl">{lbl}</p>'
#                     f'</div>',
#                     unsafe_allow_html=True,
#                 )

#         st.markdown("<br>", unsafe_allow_html=True)

#         # Downloads
#         dl1, dl2, dl3 = st.columns(3)
#         with dl1:
#             st.download_button(
#                 "⬇️ Download CSV",
#                 df.to_csv(index=False).encode("utf-8"),
#                 "screener_filtered.csv", "text/csv",
#                 use_container_width=True,
#             )
#         with dl2:
#             if st.session_state.excel1:
#                 st.download_button(
#                     "⬇️ Excel 1 — Filtered Summary",
#                     st.session_state.excel1,
#                     "screener_filtered.xlsx",
#                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                     use_container_width=True,
#                 )
#         with dl3:
#             if st.session_state.excel2:
#                 st.download_button(
#                     "⬇️ Excel 2 — Deep Company Analysis",
#                     st.session_state.excel2,
#                     "screener_detailed.xlsx",
#                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                     use_container_width=True,
#                 )

#         st.markdown("---")

#         # Searchable table
#         search = st.text_input("🔍 Search company", "")
#         fdf    = (
#             df[df["Company"].str.contains(search, case=False, na=False)]
#             if search else df.copy()
#         )
#         st.caption(f"Showing {len(fdf)} of {len(df)} companies")

#         display       = fdf.copy()
#         display["Link"] = display["Link"].apply(
#             lambda u: f'<a href="{u}" target="_blank">🔗 View</a>' if pd.notna(u) else ""
#         )

#         def color_growth(val):
#             try:
#                 v = float(val)
#                 if v >= 20:  return "background-color:#C6EFCE;color:#276221;font-weight:600"
#                 elif v >= 0: return "background-color:#FFEB9C;color:#9C5700"
#                 return "background-color:#FFC7CE;color:#9C0006"
#             except Exception:
#                 return ""

#         num_fmt     = {c: "{:,.1f}" for c in display.columns if any(q in c for q in quarters)}
#         num_fmt.update({c: "{:.2f}%" for c in display.columns if "Growth %" in c})
#         growth_cols = [c for c in display.columns if "Growth %" in c]
#         styled      = (
#             display.style
#             .applymap(color_growth, subset=growth_cols)
#             .format(num_fmt, na_rep="—")
#         )
#         st.write(styled.to_html(escape=False), unsafe_allow_html=True)

#         # Charts
#         st.markdown("---")
#         st.markdown("### 📊 Top Companies by Growth")
#         top_n = min(15, len(fdf))
#         ch1, ch2 = st.columns(2)
#         with ch1:
#             st.bar_chart(
#                 fdf.nlargest(top_n, "Sales Growth %")
#                    .set_index("Company")["Sales Growth %"]
#             )
#             st.caption(f"Top {top_n} by Sales Growth %")
#         with ch2:
#             st.bar_chart(
#                 fdf.nlargest(top_n, "Profit Growth %")
#                    .set_index("Company")["Profit Growth %"]
#             )
#             st.caption(f"Top {top_n} by Profit Growth %")

#         try:
#             import plotly.express as px
#             st.markdown("### 🔵 Sales Growth vs Profit Growth")
#             fig = px.scatter(
#                 fdf, x="Sales Growth %", y="Profit Growth %",
#                 hover_name="Company", color="Profit Growth %",
#                 color_continuous_scale="Blues", template="plotly_white",
#             )
#             fig.update_traces(marker=dict(size=9, opacity=0.8))
#             fig.update_layout(height=400, margin=dict(l=20, r=20, t=20, b=20))
#             st.plotly_chart(fig, use_container_width=True)
#         except ImportError:
#             pass

# else:
#     st.info("👆 Enter credentials, configure options, and click **Start Scraping**.")
#     with st.expander("ℹ️ How it works & what you get"):
#         st.markdown("""
# **Step 1 — Scrape & Filter**
# - Logs into Screener.in, fetches `/results/latest/` across all pages
# - **Quarters auto-detected from live HTML** — works even when Screener changes the dates
# - Keeps companies where Sales AND Net Profit are both strictly increasing Q1 → Q2 → Q3

# **Excel 1 — Filtered Summary** (`screener_filtered.xlsx`)
# - Dynamic quarter columns (whatever Screener currently shows)
# - Sales & Net Profit per quarter + growth %
# - Color-coded growth cells (green ≥ 20%, yellow ≥ 0%, red < 0%)
# - Summary stats sheet | Hyperlinked Screener links

# **Excel 2 — Deep Company Analysis** (`screener_detailed.xlsx`)
# - Visits each qualifying company's Screener page individually
# - **Sheet 1 — Key Ratios:** ROE, ROCE, P/BV, P/E, OPM, Interest Cover, Debt/Equity, Promoter Holding, Market Cap
# - **Sheet 2 — Quarterly P&L:** Last 6 quarters of Sales, Net Profit, EBIDT, EPS, OPM %
# - **Sheet 3 — Annual P&L:** Last 3 years of Sales, Net Profit, OPM %
# - ROE/ROCE/OPM cells auto-colored green (above threshold) or red

# **Tips**
# - Test with Max pages = 3 first
# - Keep delay ≥ 3s to avoid 429 errors
# - Deep detail adds ~1–2s per company
#         """)


        
################################################## Third ####################################################

"""
app.py — Streamlit UI for the Screener Growth Filter.
All scraping / parsing / Excel logic lives in scraper.py.
"""

import time
import random

import pandas as pd
import streamlit as st

from scraper import (
    login,
    run_full_scrape,
    scrape_company_detail,
    build_filtered_excel,
    build_detailed_excel,
)

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Screener Growth Filter", page_icon="📈", layout="wide")

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@500&display=swap');
  html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
  .block-container { padding-top: 1.5rem; max-width: 1400px; }

  .kpi-card {
    background: #0F172A; border-radius: 12px; padding: 20px 24px;
    color: white; border: 1px solid #1E293B; margin-bottom: 8px;
  }
  .kpi-card .val {
    font-size: 1.9rem; font-weight: 700; margin: 0;
    font-family: 'DM Mono', monospace; color: #38BDF8;
  }
  .kpi-card .lbl {
    font-size: 0.75rem; text-transform: uppercase;
    letter-spacing: 0.08em; opacity: 0.55; margin: 4px 0 0;
  }
  .criteria-box {
    background: #F0F9FF; border-left: 4px solid #0EA5E9;
    padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.92rem;
  }
  .warn-box {
    background: #FFF7ED; border-left: 4px solid #F97316;
    padding: 10px 16px; border-radius: 6px; margin-bottom: 1rem; font-size: 0.88rem;
  }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📈 Screener Quarterly Growth Filter")
st.markdown("""
<div class="criteria-box">
  <b>Filter:</b> Companies where <b>Sales</b> AND <b>Net Profit</b> are strictly increasing
  across the 3 most recent quarters — <b>quarters are auto-detected live from Screener.in HTML,
  never hardcoded</b>. Produces two formatted Excel files.
</div>
""", unsafe_allow_html=True)

# ── Credentials ────────────────────────────────────────────────────────────────
with st.expander("🔐 Screener.in Login  (required)", expanded=True):
    st.markdown("""<div class="warn-box">
      Login required — authenticated sessions receive full data and avoid 429 rate limits.
    </div>""", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        username = st.text_input("Email / Username", placeholder="you@example.com")
    with c2:
        password = st.text_input("Password", type="password")

# ── Options ────────────────────────────────────────────────────────────────────
col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    max_pages = st.number_input("Max pages (0 = all)", min_value=0, max_value=200, value=0)
with col_b:
    delay = st.slider("Delay between pages (sec)", 1.0, 6.0, 3.0, 0.5)
with col_c:
    fetch_detail = st.checkbox(
        "Fetch deep detail per company", value=True,
        help=(
            "Visits each company page — extracts ROE, ROCE, P/BV, PE/EPS, "
            "OPM (Q & P&L), Interest (Borrowings/Reserves), Promoter holding trend"
        ),
    )
with col_d:
    detail_delay = st.slider("Detail fetch delay (sec)", 0.5, 4.0, 1.5, 0.5)

st.markdown("<br>", unsafe_allow_html=True)
start = st.button("🚀 Start Scraping", type="primary")

# ── Session-state init ─────────────────────────────────────────────────────────
for key in ("df_result", "quarters", "excel1", "excel2"):
    if key not in st.session_state:
        st.session_state[key] = None

# ── Scrape trigger ─────────────────────────────────────────────────────────────
if start:
    if not username or not password:
        st.error("Please enter your Screener.in credentials.")
    else:
        progress_bar = st.progress(0, text="Starting...")
        status_text  = st.empty()

        try:
            status_text.text("🔐 Logging in...")
            session = login(username, password)
            status_text.text("✅ Logged in.")

            def on_page_start(page, total, n):
                status_text.text(f"🌐 Page {page}/{total} — {n} companies so far...")

            def on_page_done(page, total, n):
                progress_bar.progress(page / max(total, 1), text=f"Page {page} / {total}")

            df, quarters = run_full_scrape(
                session, max_pages, delay,
                on_page_start=on_page_start,
                on_page_done=on_page_done,
            )
            st.session_state.df_result = df
            st.session_state.quarters  = quarters

            status_text.text("📊 Building Excel 1 — Filtered Summary...")
            st.session_state.excel1 = build_filtered_excel(df, quarters)

            if fetch_detail and not df.empty:
                detail_records = []
                total_d = len(df)
                dp = st.progress(0, text="Fetching company details...")

                for i, row in enumerate(df.itertuples(), 1):
                    dp.progress(i / total_d, text=f"Detail {i}/{total_d}: {row.Company}")
                    status_text.text(f"🔍 Fetching: {row.Company}")

                    # Pass Profit Growth % as PAT Jumped CR
                    pat_growth = getattr(row, "Profit Growth %", None)

                    rec = scrape_company_detail(
                        session,
                        row.Company,
                        row.Link,
                        pat_growth=pat_growth,
                    )
                    detail_records.append(rec)
                    time.sleep(detail_delay * random.uniform(0.7, 1.3))

                dp.progress(1.0, text="Details done!")
                status_text.text("📊 Building Excel 2 — Deep Analysis...")
                st.session_state.excel2 = build_detailed_excel(detail_records)
            else:
                st.session_state.excel2 = None

            progress_bar.progress(1.0, text="All done!")
            status_text.success(
                f"✅ {len(df)} qualifying companies. Download your Excel files below."
            )

        except Exception as e:
            progress_bar.empty()
            status_text.error(f"❌ {e}")
            import traceback
            st.code(traceback.format_exc())

# ── Results display ────────────────────────────────────────────────────────────
df       = st.session_state.df_result
quarters = st.session_state.quarters or []

if df is not None:
    if df.empty:
        st.warning("No companies matched the filter criteria.")
    else:
        st.markdown("---")

        if quarters:
            st.info(f"**Quarters (auto-detected):** `{'  →  '.join(quarters)}`")

        # KPI cards
        avg_sg = df["Sales Growth %"].mean()  if "Sales Growth %" in df.columns else 0
        avg_pg = df["Profit Growth %"].mean() if "Profit Growth %" in df.columns else 0
        top_co = df.nlargest(1, "Profit Growth %")["Company"].values[0] if len(df) else "—"
        top_pg = df["Profit Growth %"].max()  if "Profit Growth %" in df.columns else 0

        for col, (val, lbl) in zip(st.columns(4), [
            (len(df),           "Qualifying Companies"),
            (f"{avg_sg:.1f}%",  "Avg Sales Growth"),
            (f"{avg_pg:.1f}%",  "Avg Profit Growth"),
            (f"{top_pg:.1f}%",  f"Best: {top_co}"),
        ]):
            with col:
                st.markdown(
                    f'<div class="kpi-card">'
                    f'<p class="val">{val}</p>'
                    f'<p class="lbl">{lbl}</p>'
                    f'</div>',
                    unsafe_allow_html=True,
                )

        st.markdown("<br>", unsafe_allow_html=True)

        # Downloads
        dl1, dl2, dl3 = st.columns(3)
        with dl1:
            st.download_button(
                "⬇️ Download CSV",
                df.to_csv(index=False).encode("utf-8"),
                "screener_filtered.csv", "text/csv",
                use_container_width=True,
            )
        with dl2:
            if st.session_state.excel1:
                st.download_button(
                    "⬇️ Excel 1 — Filtered Summary",
                    st.session_state.excel1,
                    "screener_filtered.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with dl3:
            if st.session_state.excel2:
                st.download_button(
                    "⬇️ Excel 2 — Deep Analysis",
                    st.session_state.excel2,
                    "screener_detailed.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        st.markdown("---")

        # Searchable table
        search = st.text_input("🔍 Search company", "")
        fdf    = (
            df[df["Company"].str.contains(search, case=False, na=False)]
            if search else df.copy()
        )
        st.caption(f"Showing {len(fdf)} of {len(df)} companies")

        display       = fdf.copy()
        display["Link"] = display["Link"].apply(
            lambda u: f'<a href="{u}" target="_blank">🔗 View</a>' if pd.notna(u) else ""
        )

        def color_growth(val):
            try:
                v = float(val)
                if v >= 20:  return "background-color:#C6EFCE;color:#276221;font-weight:600"
                elif v >= 0: return "background-color:#FFEB9C;color:#9C5700"
                return "background-color:#FFC7CE;color:#9C0006"
            except Exception:
                return ""

        num_fmt     = {c: "{:,.1f}" for c in display.columns if any(q in c for q in quarters)}
        num_fmt.update({c: "{:.2f}%" for c in display.columns if "Growth %" in c})
        growth_cols = [c for c in display.columns if "Growth %" in c]
        styled      = (
            display.style
            .applymap(color_growth, subset=growth_cols)
            .format(num_fmt, na_rep="—")
        )
        st.write(styled.to_html(escape=False), unsafe_allow_html=True)

        # Charts
        st.markdown("---")
        st.markdown("### 📊 Top Companies by Growth")
        top_n = min(15, len(fdf))
        ch1, ch2 = st.columns(2)
        with ch1:
            st.bar_chart(
                fdf.nlargest(top_n, "Sales Growth %")
                   .set_index("Company")["Sales Growth %"]
            )
            st.caption(f"Top {top_n} by Sales Growth %")
        with ch2:
            st.bar_chart(
                fdf.nlargest(top_n, "Profit Growth %")
                   .set_index("Company")["Profit Growth %"]
            )
            st.caption(f"Top {top_n} by Profit Growth %")

        try:
            import plotly.express as px
            st.markdown("### 🔵 Sales Growth vs Profit Growth")
            fig = px.scatter(
                fdf, x="Sales Growth %", y="Profit Growth %",
                hover_name="Company", color="Profit Growth %",
                color_continuous_scale="Blues", template="plotly_white",
            )
            fig.update_traces(marker=dict(size=9, opacity=0.8))
            fig.update_layout(height=400, margin=dict(l=20, r=20, t=20, b=20))
            st.plotly_chart(fig, use_container_width=True)
        except ImportError:
            pass

else:
    st.info("👆 Enter credentials, configure options, and click **Start Scraping**.")
    with st.expander("ℹ️ How it works & what you get"):
        st.markdown("""
**Step 1 — Scrape & Filter**
- Logs into Screener.in, fetches `/results/latest/` across all pages
- Quarters auto-detected from live HTML — never hardcoded
- Keeps companies where Sales AND Net Profit are both strictly increasing Q1 → Q2 → Q3

**Excel 1 — Filtered Summary** (`screener_filtered.xlsx`)
- Sales & Net Profit per quarter + growth % (colour-coded)
- Summary stats sheet | Hyperlinked Screener links

**Excel 2 — Deep Analysis** (`screener_detailed.xlsx`)
Visits each qualifying company individually and extracts:

| Column | Source | Formula |
|---|---|---|
| ROE | Key ratios | direct |
| ROCE | Key ratios | direct |
| P/BV | Key ratios | Current Price ÷ Book Value |
| PE/EPS | Key ratios + Quarterly P&L | Stock P/E ÷ latest quarterly EPS |
| PAT Jumped CR | Filter result | Profit Growth % |
| OPM (Quarterly) | Quarterly P&L | latest OPM % |
| OPM (P&L) | Annual P&L | latest OPM % |
| Interest | Balance Sheet | (Borrowings ÷ Reserves) × 100 |
| Promoter Latest % | Shareholding Pattern | latest promoter % |
| Promoter Prev % | Shareholding Pattern | second-last promoter % |
| Promoter Δ | Computed | Latest − Prev (green if +, red if −) |

**Tips**
- Test with Max pages = 3 first
- Keep delay ≥ 3s to avoid 429 errors
- Deep detail adds ~1–2s per company
        """)

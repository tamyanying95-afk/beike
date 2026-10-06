"""
把 ke_counts.json 转成 Excel + HTML 报告
按城市层级（一线 / 二线 / 三线）分组，展示：城市、二手挂牌套数、可点击网址。
没有数据的城市保留在列表中，挂牌套数记为 NA。

用法:
    python build_report.py [ke_counts.json 路径]

输出:
    output/贝壳二手挂牌套数_<日期>.xlsx
    output/贝壳二手挂牌套数_<日期>.html
"""

import json
import sys
from datetime import datetime
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Asia/Shanghai")
except Exception:
    TZ = None


def now_cn():
    return datetime.now(TZ) if TZ else datetime.now()
from xml.sax.saxutils import escape

try:
    HERE = Path(__file__).parent
except NameError:          # Jupyter cell 里没有 __file__
    HERE = Path.cwd()

OUT_DIR = HERE / "output"
OUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------- 城市清单与分组
# 顺序 = 报告中的展示顺序
CITY_SLUGS = {
    "北京": "bj",       "上海": "sh",       "深圳": "sz",       "广州": "gz",
    "天津": "tj",       "重庆": "cq",       "杭州": "hz",       "南京": "nj",
    "武汉": "wh",       "成都": "cd",       "苏州": "su",       "大连": "dl",
    "厦门": "xm",       "西安": "xa",       "长沙": "cs",       "宁波": "nb",
    "福州": "fz",       "沈阳": "sy",       "青岛": "qd",       "济南": "jn",
    "南昌": "nc",       "合肥": "hf",       "石家庄": "sjz",    "长春": "cc",
    "南宁": "nn",       "贵阳": "gy",       "昆明": "km",       "海口": "hk",
    "郑州": "zz",       "无锡": "wx",       "温州": "wz",       "三亚": "san",
    "包头": "baotou",   "东莞": "dg",       "惠州": "hui",      "中山": "zs",
    "徐州": "xz",       "芜湖": "wuhu",     "江阴": "jy",       "潍坊": "wf",
    "泸州": "luzhou",   "绍兴": "sx",       "嘉兴": "jx",       "镇江": "zj",
    "佛山": "fs",       "泰安": "ta",       "湖州": "huzhou",   "淮安": "ha",
    "常熟": "changshu", "张家口": "zjk",    "吉林": "jl",       "海门": "haimen",
}

# 一线（4）：北上广深
TIER1 = {"北京", "上海", "深圳", "广州"}

# 二线（32）：省会26 + 直辖市2（重庆/天津）+ 计划单列市4（大连/宁波/青岛/厦门）
TIER2 = {
    "长春", "长沙", "成都", "福州", "贵阳", "海口", "杭州", "哈尔滨",
    "合肥", "呼和浩特", "济南", "昆明", "兰州", "南昌", "南京", "南宁",
    "沈阳", "石家庄", "苏州", "太原", "乌鲁木齐", "武汉", "西安", "西宁",
    "银川", "郑州",
    "重庆", "天津",
    "大连", "宁波", "青岛", "厦门",
}

TIERS = [(1, "一线城市"), (2, "二线城市"), (3, "三线城市")]

MISSING_NOTE = "该城市无公开挂牌数据"

# 静态网页无法自行运行爬虫，「立即刷新」按钮指向 GitHub Actions 的手动触发页
DISPATCH_URL = "https://github.com/tamyanying95-afk/beike/actions/workflows/refresh.yml"


def load(json_path):
    d = json.loads(Path(json_path).read_text(encoding="utf-8"))
    data = d["data"]

    rows = []
    for city, slug in CITY_SLUGS.items():
        v = data.get(city) or {}
        cnt = v.get("count")
        status = v.get("status", "NOT_FETCHED")
        rows.append({
            "tier": 1 if city in TIER1 else (2 if city in TIER2 else 3),
            "city": city,
            "count": cnt,                                  # None -> NA
            "url": v.get("url") or f"https://{slug}.ke.com/ershoufang/",
            "status": status,
        })

    order = {c: i for i, c in enumerate(CITY_SLUGS)}
    rows.sort(key=lambda r: (r["tier"], order[r["city"]]))
    d["total_cities"] = len(CITY_SLUGS)
    return d, rows


# ---------------------------------------------------- openpyxl Excel 版本
def build_excel_openpyxl(d, rows, out):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "二手挂牌套数"

    tier_fill = {1: "FDECEA", 2: "E8F0FB", 3: "EFF7EF"}
    head_fill = PatternFill("solid", fgColor="2E75B6")
    sum_fill = PatternFill("solid", fgColor="DDEBF7")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ok_rows = [r for r in rows if r["count"] is not None]
    total = sum(r["count"] for r in ok_rows)

    ws.cell(1, 1, "贝壳找房 · 各城市二手房挂牌套数").font = Font(size=14, bold=True, color="1F4E79")
    ws.cell(2, 1, f"抓取时间：{d['fetched_at']}    有数据 {len(ok_rows)}/{len(rows)} 城    "
                  f"合计 {total:,} 套").font = Font(size=10, color="595959")

    hrow = 4
    for c, h in enumerate(["城市", "二手挂牌套数（套）", "网址"], 1):
        cell = ws.cell(hrow, c, h)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[hrow].height = 22

    rn = hrow + 1
    for tier, name in TIERS:
        sub = [r for r in rows if r["tier"] == tier]
        if not sub:
            continue
        ws.cell(rn, 1, f"{name}（{len(sub)} 城）")
        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=3)
        for c in range(1, 4):
            cell = ws.cell(rn, c)
            cell.fill = PatternFill("solid", fgColor=tier_fill[tier])
            cell.border = border
        ws.cell(rn, 1).font = Font(bold=True, size=11)
        ws.cell(rn, 1).alignment = Alignment(horizontal="left")
        rn += 1

        for r in sub:
            ws.cell(rn, 1, r["city"]).border = border
            if r["count"] is None:
                c2 = ws.cell(rn, 2, "NA")
                c2.font = Font(color="A6A6A6", italic=True)
                c2.alignment = Alignment(horizontal="center")
            else:
                c2 = ws.cell(rn, 2, r["count"])
                c2.number_format = "#,##0"
                c2.alignment = Alignment(horizontal="right")
            c2.border = border
            c3 = ws.cell(rn, 3, "查看")
            c3.hyperlink = r["url"]
            c3.font = Font(color="0563C1", underline="single")
            c3.alignment = Alignment(horizontal="center")
            c3.border = border
            rn += 1

    ws.cell(rn, 1, "合计（仅含有效数据）").font = Font(bold=True)
    c2 = ws.cell(rn, 2, total)
    c2.number_format = "#,##0"
    c2.alignment = Alignment(horizontal="right")
    c2.font = Font(bold=True)
    for c in range(1, 4):
        ws.cell(rn, c).fill = sum_fill
        ws.cell(rn, c).border = border

    for c, w in enumerate([18, 22, 10], 1):
        ws.column_dimensions[chr(64 + c)].width = w
    ws.freeze_panes = f"A{hrow + 1}"
    wb.save(out)
    return out


# ------------------------------------------------------- 零依赖降级版本
_WB = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
       'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
       '<sheets><sheet name="二手挂牌套数" sheetId="1" r:id="rId1"/></sheets></workbook>')

_ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              'relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

_WB_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/styles" Target="styles.xml"/></Relationships>')

_CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
       '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
       '<Default Extension="xml" ContentType="application/xml"/>'
       '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
       'spreadsheetml.sheet.main+xml"/>'
       '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-'
       'officedocument.spreadsheetml.worksheet+xml"/>'
       '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
       'spreadsheetml.styles+xml"/></Types>')

_STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
           '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>'
           '<fonts count="4">'
           '<font><sz val="11"/><name val="Calibri"/></font>'
           '<font><b/><sz val="11"/><name val="Calibri"/></font>'
           '<font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/></font>'
           '<font><i/><sz val="11"/><color rgb="FFA6A6A6"/><name val="Calibri"/></font>'
           '</fonts>'
           '<fills count="3"><fill><patternFill patternType="none"/></fill>'
           '<fill><patternFill patternType="gray125"/></fill>'
           '<fill><patternFill patternType="solid"><fgColor rgb="FFEFF3F8"/>'
           '<bgColor indexed="64"/></patternFill></fill></fills>'
           '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
           '<border><left style="thin"><color rgb="FFBFBFBF"/></left>'
           '<right style="thin"><color rgb="FFBFBFBF"/></right>'
           '<top style="thin"><color rgb="FFBFBFBF"/></top>'
           '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders>'
           '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
           '<cellXfs count="6">'
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
           '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
           '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
           '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
           '<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
           '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
           '</cellXfs></styleSheet>')

S_DEF, S_TIER, S_NUM, S_TXT, S_NA, S_TITLE = range(6)


def build_excel_stdlib(d, rows, out):
    """无 openpyxl 时的降级实现（网址以纯文本呈现，无超链接）。"""
    import zipfile

    def cell(ref, value, style):
        if value is None or value == "":
            return f'<c r="{ref}" s="{style}"/>'
        if isinstance(value, (int, float)):
            return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
        return (f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">'
                f'{escape(str(value))}</t></is></c>')

    ok_rows = [r for r in rows if r["count"] is not None]
    total = sum(r["count"] for r in ok_rows)
    body, r_i = [], 1

    body.append(f'<row r="{r_i}">{cell("A1", "贝壳找房 · 各城市二手房挂牌套数", S_TITLE)}</row>')
    r_i += 1
    sub = f"抓取时间：{d['fetched_at']}    有数据 {len(ok_rows)}/{len(rows)} 城    合计 {total:,} 套"
    body.append(f'<row r="{r_i}">{cell(f"A{r_i}", sub, S_DEF)}</row>')
    r_i += 1
    r_i += 1                                   # 空行
    hrow = r_i
    hs = "".join(cell(f"{col}{hrow}", h, S_TIER) for col, h in zip("ABC", ["城市", "二手挂牌套数（套）", "网址"]))
    body.append(f'<row r="{hrow}">{hs}</row>')
    r_i += 1

    for tier, name in TIERS:
        sub_rows = [r for r in rows if r["tier"] == tier]
        if not sub_rows:
            continue
        body.append(f'<row r="{r_i}">{cell(f"A{r_i}", f"{name}（{len(sub_rows)} 城）", S_TIER)}</row>')
        r_i += 1
        for r in sub_rows:
            if r["count"] is None:
                c2 = cell(f"B{r_i}", "NA", S_NA)
            else:
                c2 = cell(f"B{r_i}", r["count"], S_NUM)
            body.append(f'<row r="{r_i}">{cell(f"A{r_i}", r["city"], S_TXT)}'
                        f'{c2}{cell(f"C{r_i}", r["url"], S_TXT)}</row>')
            r_i += 1
    body.append(f'<row r="{r_i}">{cell(f"A{r_i}", "合计（仅含有效数据）", S_TIER)}'
                f'{cell(f"B{r_i}", total, S_NUM)}</row>')

    sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             f'<dimension ref="A1:C{r_i}"/>'
             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" '
             'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
             '<sheetFormatPr defaultRowHeight="15"/>'
             '<cols><col min="1" max="1" width="18" customWidth="1"/>'
             '<col min="2" max="2" width="22" customWidth="1"/>'
             '<col min="3" max="3" width="34" customWidth="1"/></cols>'
             f'<sheetData>{"".join(body)}</sheetData></worksheet>')

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("xl/workbook.xml", _WB)
        z.writestr("xl/_rels/workbook.xml.rels", _WB_RELS)
        z.writestr("xl/styles.xml", _STYLES)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return out


def build_excel(d, rows, out):
    try:
        build_excel_openpyxl(d, rows, out)
        return out, "openpyxl（含超链接）"
    except ImportError:
        build_excel_stdlib(d, rows, out)
        return out, "stdlib 降级（网址为纯文本）"


# ------------------------------------------------------------------- HTML
HTML_TPL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>贝壳二手挂牌套数 · {date}</title>
<style>
  :root {{
    --bg:#f7f8fa; --card:#fff; --line:#e6e8eb; --text:#1f2328;
    --muted:#6b7280; --brand:#c0392b; --brand-soft:#fdf1ef; --blue:#2e75b6;
  }}
  *{{box-sizing:border-box}}
  body{{margin:0;background:var(--bg);color:var(--text);
       font:14px/1.6 -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif}}
  .wrap{{max-width:820px;margin:0 auto;padding:32px 20px 60px}}
  h1{{font-size:24px;margin:0 0 6px}}
  .sub{{color:var(--muted);font-size:13px;margin-bottom:12px}}
  .bar{{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:18px}}
  .btn{{display:inline-block;padding:8px 16px;border-radius:8px;background:var(--brand);
       color:#fff;font-size:13px;text-decoration:none;line-height:1.2}}
  .btn:hover{{filter:brightness(1.08);text-decoration:none}}
  .tip{{font-size:12px;color:var(--muted);line-height:1.5}}
  .tip a{{font-size:12px}}
  .cards{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:22px}}
  .card{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:16px 18px}}
  .card .k{{color:var(--muted);font-size:12px}}
  .card .v{{font-size:26px;font-weight:600;margin-top:4px}}
  .card .v small{{font-size:13px;font-weight:400;color:var(--muted)}}
  .v.red{{color:var(--brand)}}
  input{{width:100%;padding:9px 12px;border:1px solid var(--line);border-radius:8px;
        background:var(--card);font-size:14px;outline:none;margin-bottom:12px}}
  input:focus{{border-color:var(--blue)}}
  table{{width:100%;border-collapse:collapse;background:var(--card);
        border:1px solid var(--line);border-radius:10px;overflow:hidden}}
  th{{background:#f0f4f8;text-align:center;font-weight:600;font-size:13px;
     padding:11px 10px;border-bottom:1px solid var(--line);white-space:nowrap}}
  td{{padding:9px 12px;border-bottom:1px solid #f0f1f3;font-size:13px}}
  td.city{{font-weight:500}}
  td.num{{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;width:150px}}
  td.link{{text-align:center;width:90px}}
  tr.city:hover td{{background:var(--brand-soft)}}
  tr.tier td{{font-weight:600;font-size:13px;padding:8px 12px;border-bottom:1px solid var(--line)}}
  tr.tier-1 td{{background:#fdecea;color:#a93226}}
  tr.tier-2 td{{background:#e8f0fb;color:#1f4e79}}
  tr.tier-3 td{{background:#eff7ef;color:#1e6b3a}}
  tr.na td.city,tr.na td.num{{color:#9aa0a6;font-weight:400}}
  tr.na td.num{{font-style:italic;text-align:center}}
  tfoot td{{background:#f0f4f8;font-weight:600;border-top:1px solid var(--line)}}
  a{{color:var(--blue);text-decoration:none;font-size:13px}}
  a:hover{{text-decoration:underline}}
  .note{{margin-top:16px;color:var(--muted);font-size:12px}}
  @media(max-width:640px){{.cards{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="wrap">
  <h1>贝壳找房 · 各城市二手房挂牌套数</h1>
  <div class="sub">数据来源：贝壳找房（ke.com）各城市二手房源列表页　·　抓取时间：{fetched_at}（北京时间）</div>

  <div class="bar">
    <a class="btn" href="{dispatch_url}" target="_blank" rel="noopener">立即刷新</a>
    <span class="tip">正常情况下<b>每周三 09:00</b>自动刷新。静态网页无法自行运行爬虫，
      临时刷新请点左侧按钮 → 在 GitHub 页面点 <b>Run workflow</b>，约 5–10 分钟后网页更新。</span>
  </div>

  <div class="cards">
    <div class="card"><div class="k">有数据城市</div><div class="v">{ok} <small>/ {total_cities} 城</small></div></div>
    <div class="card"><div class="k">挂牌总量</div><div class="v red">{total_fmt} <small>套</small></div></div>
    <div class="card"><div class="k">榜首城市</div><div class="v">{top_city} <small>{top_cnt} 套</small></div></div>
  </div>

  <input id="q" placeholder="搜索城市…">

  <table id="t">
    <thead><tr><th>城市</th><th>二手挂牌套数（套）</th><th>网址</th></tr></thead>
    <tbody>
{rows}
    </tbody>
    <tfoot><tr><td>合计（仅含有效数据）</td><td class="num">{total_fmt}</td><td></td></tr></tfoot>
  </table>
  <div class="note">NA = {missing_note}；点击查看可打开对应城市的贝壳二手房源列表页。</div>
</div>
<script>
const q=document.getElementById('q');
q.oninput=()=>{{const v=q.value.trim();
  document.querySelectorAll('#t tbody tr.city').forEach(r=>
    r.style.display=(!v||r.cells[0].innerText.includes(v))?'':'none');}};
</script>
</body></html>
"""


def build_html(d, rows, out):
    ok_rows = [r for r in rows if r["count"] is not None]
    total = sum(r["count"] for r in ok_rows)
    top = max(ok_rows, key=lambda r: r["count"]) if ok_rows else {"city": "-", "count": 0}

    lines = []
    for tier, name in TIERS:
        sub = [r for r in rows if r["tier"] == tier]
        if not sub:
            continue
        lines.append(
            f'      <tr class="tier tier-{tier}"><td colspan="3">{name}（{len(sub)} 城）</td></tr>'
        )
        for r in sub:
            if r["count"] is None:
                cnt, cls = "NA", ' class="city na"'
            else:
                cnt, cls = f"{r['count']:,}", ' class="city"'
            lines.append(
                f'      <tr{cls}><td class="city">{r["city"]}</td>'
                f'<td class="num">{cnt}</td>'
                f'<td class="link"><a href="{r["url"]}" target="_blank" rel="noopener">查看</a></td></tr>'
            )

    Path(out).write_text(HTML_TPL.format(
        date=now_cn().strftime("%Y-%m-%d"),
        fetched_at=d["fetched_at"],
        ok=len(ok_rows),
        total_cities=len(rows),
        total_fmt=f"{total:,}",
        top_city=top["city"],
        top_cnt=f"{top['count']:,}" if ok_rows else "-",
        missing_note=MISSING_NOTE,
        dispatch_url=DISPATCH_URL,
        rows="\n".join(lines),
    ), encoding="utf-8")
    return out


def main():
    cands = ([Path(sys.argv[1])] if len(sys.argv) > 1
             else [OUT_DIR / "ke_counts.json", HERE / "ke_counts.json",
                   Path.cwd() / "ke_counts.json",
                   Path.cwd() / "output" / "ke_counts.json"])
    src = next((p for p in cands if p.exists()), None)
    if src is None:
        raise SystemExit("找不到 ke_counts.json，找过：\n  "
                         + "\n  ".join(str(p) for p in cands)
                         + "\n先跑 scrape_ke_counts.py，或把路径当参数传入。")

    d, rows = load(src)
    stamp = now_cn().strftime("%Y%m%d")
    xlsx, engine = build_excel(d, rows, OUT_DIR / f"贝壳二手挂牌套数_{stamp}.xlsx")
    html = build_html(d, rows, OUT_DIR / f"贝壳二手挂牌套数_{stamp}.html")
    print("生成完成：")
    print("  Excel:", xlsx, f"（{engine}）")
    print("  HTML :", html)


if __name__ == "__main__":
    main()

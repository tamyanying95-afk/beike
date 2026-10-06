"""
贝壳/链家 二手房挂牌套数 抓取脚本
在 Jupyter 里 %run 本文件，或直接粘贴到 cell 执行。

每个城市的抓取地址形如 https://bj.ke.com/ershoufang/ （北京→bj，福州→fz，依此类推）
直连访问，不配置任何代理。

产出: <脚本目录>/output/ke_counts.json  (供 build_report.py 生成 HTML + Excel)

依赖: requests (已装则可跳过)
"""

import json
import os
import random
import re
import socket
import time
from datetime import datetime
from pathlib import Path

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------- 配置
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

# 直连，不走代理。ke.com 各城市站形如 https://bj.ke.com/ershoufang/
ROUNDS = 2                      # 对失败城市补抓几轮
SLEEP_RANGE = (1.5, 3.0)        # 每次请求之间的随机间隔(秒)

# 只补抓指定城市时填这里，例如 ["上海","重庆","武汉"]；留空则抓全部
ONLY = []

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

# Jupyter 的 cell 里没有 __file__，用当前工作目录兜底
try:
    HERE = Path(__file__).parent
except NameError:
    HERE = Path.cwd()

# 想换输出位置，直接改这一行即可
OUT_DIR = HERE / "output"
OUT_DIR.mkdir(exist_ok=True)


# ---------------------------------------------------------------- session
def create_session():
    s = requests.Session()

    # 默认直连、不配置代理。
    # 若设置了 HTTPS_PROXY / HTTP_PROXY 环境变量（GitHub Actions 里用 secrets 配），
    # 则自动走该代理——以后 IP 被挡时不用改代码。
    proxy = (os.getenv("HTTPS_PROXY") or os.getenv("https_proxy")
             or os.getenv("HTTP_PROXY") or os.getenv("http_proxy"))
    if proxy:
        s.trust_env = True
        s.proxies = {"http": proxy, "https": proxy}
        print("使用代理:", proxy.split("@")[-1])
    else:
        s.trust_env = False
        s.proxies = {}

    retry = Retry(total=2, connect=2, read=2, backoff_factor=1,
                  status_forcelist=[429, 500, 502, 503, 504],
                  raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({"User-Agent": UA})
    return s


def count_from_html(html: str):
    """从页面里挖挂牌总数：主用 count: 数字，备用 共找到<span>N</span>套。"""
    for pattern in (r'count:\s*(\d+)', r'"totalNum"\s*:\s*(\d+)', r'"total"\s*:\s*(\d+)'):
        m = re.search(pattern, html)
        if m:
            v = int(m.group(1))
            if v > 0:
                return v
    m = re.search(r'共找到\s*<span[^>]*>\s*([\d,]+)\s*</span>', html)
    if m:
        return int(m.group(1).replace(",", ""))
    return None


def fetch(session, city, slug):
    """返回 (套数 or None, 状态说明)"""
    host = f"{slug}.ke.com"
    # 域名存不存在：不存在的直接跳过，别浪费请求
    try:
        socket.gethostbyname(host)
    except socket.gaierror:
        return None, "NO_SITE", None

    url = f"https://{host}/ershoufang/"
    try:
        r = session.get(url, timeout=25)
    except Exception as e:
        return None, f"REQ_ERR:{type(e).__name__}", url

    if "hip.ke.com/captcha" in r.url:
        return None, "CAPTCHA", url
    if r.status_code != 200:
        return None, f"HTTP_{r.status_code}", url

    n = count_from_html(r.text)
    if n is None:
        return None, "NO_MATCH", url
    return n, "OK", url


# ---------------------------------------------------------------- 主流程
def main():
    session = create_session()

    # warm-up：先拿 cookie，没有 cookie 几乎必定跳验证码
    try:
        session.get("https://www.ke.com/", timeout=20)
        session.headers.update({
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Referer": "https://www.ke.com/",
        })
    except Exception as e:
        print("warm-up 失败(忽略):", e)

    pending = [(c, s) for c, s in CITY_SLUGS.items() if (not ONLY or c in ONLY)]
    results = {}
    last_status = {}

    for rnd in range(1, ROUNDS + 1):
        if not pending:
            break
        print(f"\n=== 第 {rnd} 轮，待抓 {len(pending)} 个城市 ===")
        still = []
        for city, slug in pending:
            n, status, url = fetch(session, city, slug)
            last_status[city] = status
            if n is not None:
                results[city] = {"count": n, "slug": slug, "url": url, "status": "OK"}
            elif status == "NO_SITE":
                results[city] = {"count": None, "slug": slug, "url": None, "status": "NO_SITE"}
            else:
                still.append((city, slug))
            print(f"{city:<6} {n if n is not None else status}")
            time.sleep(random.uniform(*SLEEP_RANGE))
        pending = still
        if pending:
            time.sleep(random.uniform(8, 15))   # 一轮之间歇一歇

    # 多轮都没抓到的城市也写进 JSON，避免报表里直接消失
    for city, slug in pending:
        results[city] = {"count": None, "slug": slug,
                         "url": f"https://{slug}.ke.com/ershoufang/",
                         "status": last_status.get(city, "NO_MATCH")}

    # 补抓模式：与已有 JSON 合并，不影响之前成功的城市
    out = OUT_DIR / "ke_counts.json"
    if ONLY and out.exists():
        try:
            old = json.loads(out.read_text(encoding="utf-8"))
            merged = dict(old.get("data", {}))
            merged.update(results)
            results = merged
            print(f"（已与原有 {out.name} 合并）")
        except Exception as e:
            print("合并旧 JSON 失败，直接覆盖:", e)

    payload = {
        "fetched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_cities": len(CITY_SLUGS),
        "success": sum(1 for v in results.values() if v["count"] is not None),
        "data": results,
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    ok = [c for c, v in results.items() if v["count"] is not None]
    print(f"\n完成 {len(ok)}/{len(CITY_SLUGS)} 个城市 -> {out}")
    for c, v in results.items():
        if v["count"] is None:
            print(f"  失败: {c} ({v['status']})")
    return payload


if __name__ == "__main__":
    main()

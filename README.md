# beike · 贝壳各城市二手挂牌套数

按 Tier-1（一线）/ Tier-2（二线）/ Tier-3（三线）分组，展示 52 城的贝壳二手挂牌套数，
每周自动刷新一次并发布到 GitHub Pages。

- **在线报表**：<https://tamyanying95-afk.github.io/beike/>（repo 建好后约 1 分钟生效）
- **数据来源**：贝壳找房各城市二手房源列表页 `https://{city}.ke.com/ershoufang/`
- **字段**：城市 / 二手挂牌套数（套）/ 查看网址；无公开数据的城市记为 `NA`

## 首次配置 GitHub Pages（只需一次）

仓库 **Settings → Pages → Build and deployment**：

- Source 选 `Deploy from a branch`
- Branch 选 `main`，文件夹选 `/docs`
- Save，约 1 分钟后访问 <https://tamyanying95-afk.github.io/beike/>

注：免费账号的 Pages 只对 **public** 仓库开放，所以这个 repo 保持 public。

## 每周自动刷新

`.github/workflows/refresh.yml` 每周三 **09:00（北京时间）** 触发一次：

1. `python scrape_ke_counts.py` → 抓取全部城市，产出 `output/ke_counts.json`
2. `python build_report.py` → 生成 `output/贝壳二手挂牌套数_<日期>.html` / `.xlsx`
3. 体检：若本次成功城市数少于上次的一半（或为 0），判定被反爬拦截 → **保留旧数据不发布**
4. 通过后：`docs/index.html` 更新，快照写入 `data/latest.json` 与 `history/<日期>.json`，自动 commit

也可以随时去仓库 **Actions → Refresh Beike Data → Run workflow** 手动跑一次。

## 本地跑

```bash
pip install -r requirements.txt
python scrape_ke_counts.py     # 抓取，约 5–10 分钟（每城间隔 1.5–3 秒）
python build_report.py         # 生成 HTML + Excel 到 output/
```

## 目录说明

| 路径 | 用途 |
| --- | --- |
| `scrape_ke_counts.py` | 抓取脚本（含 cookie warm-up、DNS 预检、失败补抓 2 轮） |
| `build_report.py` | 读 `ke_counts.json` → 生成 HTML + Excel |
| `docs/` | GitHub Pages 站点目录（`index.html` 即在线报表） |
| `data/latest.json` | 最近一次成功快照 |
| `history/<日期>.json` | 历史快照，可用于做趋势图 |

## 注意

贝壳对单一 IP 有频控，各城市显示的套数还受当地房源上架/去重规则影响，
跨城直接比大小意义有限，更适合看同城市的周度变化。

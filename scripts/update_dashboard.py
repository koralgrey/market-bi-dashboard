#!/usr/bin/env python3
"""Update the embedded dashboard dataset from Tmall, JD and Douyin workbooks."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import load_workbook


DATA_RE = re.compile(r"var U = (\{.*?\});\s*</script>", re.S)
JD_ALIASES = {
    "乳胶漆": "内墙乳胶漆",
    "普通水槽": "水槽",
    "毛巾架浴巾架": "毛巾架",
    "家装隔音材料": "隔音材料",
    "补墙膏": "墙面修补膏",
    "防水卷材": "防水卷材",
    "涂刷化工类辅料": "胶类",
}
DY_FOCUS = {
    ("油漆及辅料", "美缝剂"): "美缝剂",
    ("涂料辅料", "防水涂料"): "防水涂料",
    ("防护材料", "防水喷剂"): "防水喷剂",
    ("防护材料", "防水卷材/胶带"): "防水卷材",
    ("胶粘剂", "玻璃胶/防霉胶"): "玻璃胶",
    ("胶粘剂", "密封胶"): "密封胶",
    ("胶粘剂", "免钉胶"): "免钉胶",
    ("胶粘剂", "结构胶"): "结构胶",
    ("胶粘剂", "瓷砖胶"): "瓷砖胶",
    ("油漆及辅料", "地坪漆"): "地坪漆",
    ("油漆及辅料", "墙面修补膏"): "墙面修补膏",
    ("油漆及辅料", "腻子粉/膏"): "腻子",
    ("墙面材料", "内墙乳胶漆"): "内墙乳胶漆",
    ("墙面材料", "外墙乳胶漆"): "外墙乳胶漆",
    ("油漆及辅料", "金属漆"): "金属漆",
    ("油漆及辅料", "艺术漆"): "艺术漆",
    ("墙面材料", "艺术漆"): "艺术漆",
    ("水槽与龙头", "厨盆/水槽"): "水槽",
    ("水槽与龙头", "面盆龙头"): "面盆龙头",
    ("水槽与龙头", "厨房龙头"): "厨房龙头",
    ("水槽与龙头", "洗衣机龙头"): "洗衣机龙头",
    ("水槽与龙头", "浴缸淋浴龙头"): "浴缸淋浴龙头",
    ("全屋智能", "智能龙头"): "智能龙头",
    ("卫浴五金", "地漏"): "地漏",
    ("卫浴五金", "浴巾架/毛巾架"): "毛巾架",
    ("卫浴五金", "角阀"): "角阀",
    ("卫浴建材", "浴室柜组合"): "浴室柜",
    ("卫浴建材", "花洒"): "花洒",
    ("全屋智能", "智能花洒"): "智能花洒",
    ("全屋智能", "智能门锁"): "智能门锁",
    ("全屋智能", "智能马桶"): "智能马桶",
}


def number(value):
    return float(value or 0)


def metric_series(monthly, months):
    out = {}
    for index, month in enumerate(months):
        sales = number(monthly.get(month))
        previous = number(monthly.get(months[index - 1])) if index else 0
        prior_year = number(monthly.get(f"{int(month[:4]) - 1}{month[4:]}"))
        prior_previous = number(monthly.get(months[index - 2])) if index > 1 else 0
        mom = (sales - previous) / previous * 100 if previous else None
        yoy = (sales - prior_year) / prior_year * 100 if prior_year else None
        prior_mom = (previous - prior_previous) / prior_previous * 100 if prior_previous else None
        momentum = ((mom - prior_mom) / abs(prior_mom) * 100
                    if mom is not None and prior_mom not in (None, 0) else None)
        out[month] = {
            "sales": sales,
            "mom": round(mom, 2) if mom is not None else None,
            "yoy": round(yoy, 2) if yoy is not None else None,
            "momentum": round(momentum, 2) if momentum is not None else None,
        }
    return out


def strength(metrics, latest):
    item = metrics.get(latest, {})
    mom, yoy = item.get("mom"), item.get("yoy")
    level = "强势" if (mom or 0) > 0 and (yoy or 0) > 0 else "成长" if (mom or 0) > 0 else "稳定" if (yoy or 0) > 0 else "弱势"
    return {"level": level, "mom": mom or 0, "yoy": yoy or 0}


def update_tmall(data, path):
    wb = load_workbook(path, read_only=True, data_only=True)
    rows, months = [], []
    for sheet_name in ("基础建材", "家装主材"):
        ws = wb[sheet_name]
        headers = [ws.cell(4, column).value for column in range(1, ws.max_column + 1)]
        month_cols = [(index, value) for index, value in enumerate(headers) if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}", value)]
        months = sorted(set(months).union(value for _, value in month_cols))
        for values in ws.iter_rows(min_row=5, max_col=len(headers), values_only=True):
            if not values[4]:
                continue
            monthly = {month: number(values[index]) for index, month in month_cols}
            rows.append({"一级类目": str(values[1]), "二级类目": str(values[2]), "三级类目": str(values[3]), "类目ID": str(values[4]), "monthly": monthly})
    months.sort()
    if not months or not rows:
        raise ValueError("天猫源表缺少有效类目或月份")
    identities = [(row["一级类目"], row["类目ID"]) for row in rows]
    if len(identities) != len(set(identities)):
        raise ValueError("天猫源表存在重复类目ID，停止更新")
    # Reconcile each monthly detail total to the workbook's independent summary.
    summary = wb["汇总"]
    header = next(summary.iter_rows(min_row=4, max_row=4, values_only=True))
    summary_months = {str(value): index for index, value in enumerate(header) if str(value) in months}
    summary_rows = {values[1]: values for values in summary.iter_rows(min_row=5, values_only=True)
                    if len(values) > 1 and values[1] in ("基础建材", "家装主材")}
    for sheet_name in ("基础建材", "家装主材"):
        if sheet_name not in summary_rows:
            raise ValueError(f"汇总页缺少{sheet_name}")
        for month in months:
            actual = sum(row["monthly"][month] for row in rows if row["一级类目"] == sheet_name)
            expected = number(summary_rows[sheet_name][summary_months[month]])
            if abs(actual - expected) > 0.01:
                raise ValueError(f"{sheet_name} {month} 明细与汇总不符：{actual} != {expected}")
    l3 = []
    for row in rows:
        metrics = metric_series(row.pop("monthly"), months)
        path_text = " > ".join([row["一级类目"], row["二级类目"], row["三级类目"]])
        l3.append({**row, "_navL1": row["一级类目"], "_navL2": row["二级类目"], "_navL3": row["三级类目"], "类目路径": path_text,
                   "合计": sum(x["sales"] for x in metrics.values()), "metrics": metrics, "strength": strength(metrics, months[-1])})
    groups = defaultdict(lambda: defaultdict(float))
    for row in l3:
        for month, value in row["metrics"].items():
            groups[(row["一级类目"], row["二级类目"])][month] += value["sales"]
    l2 = []
    for (level1, level2), monthly in groups.items():
        metrics = metric_series(monthly, months)
        l2.append({"一级类目": level1, "二级类目": level2, "合计": sum(monthly.values()), "metrics": metrics})
    totals = {month: sum(row["metrics"][month]["sales"] for row in l3) for month in months}
    data["months"] = months
    data["tmall"] = {"l3_data": l3, "l2_data": l2, "summary": {"category_count": len(l3), "total_sales": sum(totals.values()), "monthly_totals": totals}}
    wb.close()


def add_tmall_source_notice(html, data):
    """Keep aggregate category coverage separate from historical brand samples."""
    meta = data.get("source_metadata", {}).get("tmall", {})
    if not meta:
        return html
    notice = (
        '<div id="tmallSourceNotice" style="margin:12px 0;padding:14px 18px;'
        'background:#fff8e6;border:1px solid #ead6a0;border-radius:8px;line-height:1.8">'
        f'天猫类目销售额：{meta["first_month"]}—{meta["latest_month"]}，单位：元。'
        f'品牌与雨虹市占仍为历史样本（截至{meta["brand_latest_month"]}），'
        '尚未按当前品牌规则复核。本次源表无品牌、店铺和商品明细，未更新品牌市占。'
        f'已核对资料库当前身份规则{meta["brand_rule_version"]}；'
        '类目全量销售额与品牌排名采集样本为不同范围，不能混作市占分母。'
        '</div>'
    )
    html = re.sub(r'<div id="tmallSourceNotice".*?</div>', '', html, flags=re.S)
    marker = '<section id="multiTmall" class="platform-section">'
    html = html.replace(marker, marker + '\n' + notice, 1)
    status = (
        f'数据截至：天猫类目 {meta["latest_month"]}；'
        f'京东 {data["jd"]["months"][-1]}；抖音 {data["dy"]["months"][-1]}。'
        f'天猫品牌历史样本截至 {meta["brand_latest_month"]}，待当前规则复核。'
    )
    html = re.sub(r'<div id="platformDataCoverage".*?</div>', '', html, flags=re.S)
    header_end = '<span class="time">'
    start = html.find(header_end)
    if start >= 0:
        end = html.find('</span>', start) + len('</span>')
        html = html[:end] + '<div id="platformDataCoverage" style="font-size:12px;color:#64748b;margin-top:6px">' + status + '</div>' + html[end:]
    return html


def market_payload(records, months):
    categories = defaultdict(list)
    for record in records:
        categories[record["category"]].append(record)
    market = {}
    for category, items in categories.items():
        monthly_total = {month: 0 for month in months}
        brand_monthly = defaultdict(lambda: {month: 0 for month in months})
        for item in items:
            monthly_total[item["month"]] += item["sales"]
            brand_monthly[item["brand"]][item["month"]] += item["sales"]
        brands = []
        ranks = {month: {} for month in months}
        for month in months:
            ordered = sorted(((name, values[month]) for name, values in brand_monthly.items() if values[month] > 0), key=lambda pair: -pair[1])
            ranks[month] = {name: index + 1 for index, (name, _) in enumerate(ordered)}
        for name, values in brand_monthly.items():
            brands.append({"name": name, "total": sum(values.values()), "monthly": values,
                           "share": {month: round(values[month] / monthly_total[month] * 100, 2) if monthly_total[month] else 0 for month in months},
                           "rank": {month: ranks[month].get(name) for month in months}})
        brands.sort(key=lambda item: -item["total"])
        market[category] = {"months": months, "monthly_total": monthly_total, "brands": brands}
    return market


def update_jd(data, path):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["数据源"]
    rows, months = [], set()
    for values in ws.iter_rows(min_row=2, min_col=1, max_col=19, values_only=True):
        raw, _, _, sku, _, brand, shop, _, year, month, period, sales, fake = values[:13]
        if period != "月" or fake != "否" or not raw or not year or not month:
            continue
        month_key = f"{int(year):04d}-{int(month):02d}"
        months.add(month_key)
        rows.append({"category": JD_ALIASES.get(str(raw), str(raw)), "sku": str(sku), "brand": str(brand or "未知"), "shop": str(shop or "未知"), "month": month_key, "sales": number(sales)})
    months = sorted(months)
    all_market = market_payload(rows, months)
    self_market = market_payload([row for row in rows if row["shop"] == "自营"], months)
    pop_market = market_payload([row for row in rows if row["shop"] == "POP"], months)
    data["jd"] = {"months": months, "latest_month": months[-1], "market": all_market, "market_all": all_market,
                  "market_self": self_market, "market_pop": pop_market, "focus_market": all_market}


def update_dy(data, path):
    wb = load_workbook(path, read_only=True, data_only=True)
    july = {}
    for sheet_name in wb.sheetnames[1:]:
        ws = wb[sheet_name]
        for name, sales, _ in ws.iter_rows(min_row=4, min_col=1, max_col=3, values_only=True):
            if name:
                july[(sheet_name, str(name))] = number(sales)
    months = sorted(set(data["dy"].get("months", []) + ["2026-07"]))
    for row in data["dy"]["l3_data"]:
        key = (row["一级类目"], row["二级类目"])
        row.setdefault("metrics", {})["2026-07"] = {"sales": july.get(key, 0), "mom": None, "yoy": None, "momentum": None}
        monthly = {month: number(row["metrics"].get(month, {}).get("sales")) for month in months}
        row["metrics"] = metric_series(monthly, months)
        row["合计"] = sum(monthly.values())
        row["strength"] = strength(row["metrics"], months[-1])
        row["类目ID"] = row.get("类目ID") or f"DY|{row['一级类目']}|{row['二级类目']}"
    data["dy"]["l2_data"] = [{key: value for key, value in row.items() if key in ("一级类目", "二级类目", "合计", "metrics")} for row in data["dy"]["l3_data"]]
    totals = {month: sum(number(row["metrics"].get(month, {}).get("sales")) for row in data["dy"]["l3_data"]) for month in months}
    data["dy"]["months"] = months
    data["dy"]["summary"] = {"category_count": len(data["dy"]["l3_data"]), "total_sales": sum(totals.values()), "monthly_totals": totals}
    combined = {}
    for key, sales in july.items():
        focus = DY_FOCUS.get(key)
        if focus:
            combined[focus] = combined.get(focus, 0) + sales
    analysis = data["dy"].setdefault("product_analysis", {})
    target = analysis.setdefault("main", {})
    for focus, sales in combined.items():
        item = target.setdefault(focus, {"months": [], "monthly_total": {}, "brands": []})
        item["months"] = sorted(set(item.get("months", []) + ["2026-07"]))
        item.setdefault("monthly_total", {})["2026-07"] = sales
        brands = [brand for brand in item.setdefault("brands", []) if brand.get("name") != "未提供品牌明细"]
        brands.append({"name": "未提供品牌明细", "total": sales, "monthly": {"2026-07": sales}, "share": {"2026-07": 100}, "rank": {"2026-07": 1}})
        item["brands"] = brands


def inject_jd_selector(html):
    html = re.sub(r"生成时间: [^<]+ \| 天猫·京东·抖音", f"生成时间: {datetime.now().isoformat(timespec='seconds')} | 天猫·京东·抖音", html, count=1)
    html = html.replace("'乳胶漆=内墙乳胶漆=外墙乳胶漆'", "'内墙乳胶漆','外墙乳胶漆'")
    if 'id="jdShopType"' not in html:
        html = html.replace('<div class="platform-heading">🌐 多平台-京东 <span>', '<div class="platform-heading">🌐 多平台-京东 <select id="jdShopType" onchange="setJdShopType(this.value)" style="margin-left:16px;padding:7px 12px;border:1px solid #d1d5db;border-radius:8px"><option value="all">全部</option><option value="self">自营</option><option value="pop">POP</option></select> <span>')
    marker = "function renderMultiPlatform(platform){"
    if "function setJdShopType" not in html:
        code = "var jdShopType='all';\nfunction activeJdMarket(){return jdShopType==='self'?(JD.market_self||{}):jdShopType==='pop'?(JD.market_pop||{}):(JD.market_all||JD.market||{});}\nfunction setJdShopType(value){jdShopType=value;renderMultiPlatform('jd');}\nwindow.setJdShopType=setJdShopType;\n\n"
        html = html.replace(marker, code + marker)
        html = html.replace("renderCategorySalesYoY('cJdTrend',JD.market||{},1e4,'万元','jd');", "renderCategorySalesYoY('cJdTrend',activeJdMarket(),1e4,'万元','jd');")
        html = html.replace("renderCategoryBrandRank('cJdBrand',JD.market||{},'jd',1e4,'万元');", "renderCategoryBrandRank('cJdBrand',activeJdMarket(),'jd',1e4,'万元');")
        html = html.replace("renderYuhongShareRank('cYuhong','yuhongHighlight',JD.market||{},'京东','jd',1e4,'万元');", "renderYuhongShareRank('cYuhong','yuhongHighlight',activeJdMarket(),'京东','jd',1e4,'万元');")
    return html


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tmall")
    parser.add_argument("--jd")
    parser.add_argument("--dy")
    parser.add_argument("--html", default="index.html")
    parser.add_argument("--rules-db", help="只读查询资料库当前品牌及市占规则版本")
    parser.add_argument("--audit-report", help="保存汇总校验与历史修订审计JSON")
    args = parser.parse_args()
    target = Path(args.html)
    html = target.read_text(encoding="utf-8")
    match = DATA_RE.search(html)
    if not match:
        raise SystemExit("Embedded dashboard data was not found")
    data = json.loads(match.group(1))
    if not any((args.tmall, args.jd, args.dy)):
        parser.error("至少提供一个平台的源文件")
    audit = {"updated_platforms": [], "source_fields": "category-level aggregates; no brand/store/product detail"}
    if args.tmall:
        old = data["tmall"]["summary"]["monthly_totals"].copy()
        update_tmall(data, args.tmall)
        brand_months = sorted({month for item in data.get("tmall_market", {}).values() for month in item.get("months", [])})
        rule_version, market_rule = "未查询", "未查询"
        if args.rules_db:
            with sqlite3.connect(Path(args.rules_db).resolve().as_uri() + "?mode=ro", uri=True) as db:
                versions = db.execute("SELECT version_code FROM brand_store_mapping_versions WHERE status='当前有效'").fetchall()
                if len(versions) != 1:
                    raise ValueError("资料库必须只有一个当前品牌规则版本")
                rule_version = versions[0][0]
                market_rule = db.execute("SELECT current_version FROM market_share_governance").fetchone()[0]
        meta = {"filename": Path(args.tmall).name, "sha256": hashlib.sha256(Path(args.tmall).read_bytes()).hexdigest(),
                "first_month": data["months"][0], "latest_month": data["months"][-1], "unit": "元",
                "sheets": ["基础建材", "家装主材"], "brand_detail_available": False,
                "brand_latest_month": brand_months[-1] if brand_months else "无数据",
                "brand_rule_version": rule_version, "market_rule_version": market_rule,
                "brand_status": "历史样本未按当前身份规则重新核验；不作为最新正式品牌市占"}
        data.setdefault("source_metadata", {})["tmall"] = meta
        totals = data["tmall"]["summary"]["monthly_totals"]
        audit.update({"tmall": meta, "category_count": len(data["tmall"]["l3_data"]), "monthly_totals_yuan": totals,
                      "summary_reconciliation": "passed_all_months_both_sheets",
                      "historical_revisions": [{"month": month, "old_yuan": value, "new_yuan": totals[month], "delta_yuan": totals[month] - value}
                                               for month, value in old.items() if month in totals and abs(totals[month] - value) > 0.01]})
        audit["updated_platforms"].append("tmall")
    if args.jd:
        update_jd(data, args.jd)
        audit["updated_platforms"].append("jd")
    if args.dy:
        update_dy(data, args.dy)
        audit["updated_platforms"].append("dy")
    data["generated_at"] = datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = html[:match.start(1)] + payload + html[match.end(1):]
    html = inject_jd_selector(html)
    html = add_tmall_source_notice(html, data)
    target.write_text(html, encoding="utf-8")
    if args.audit_report:
        report = Path(args.audit_report)
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"generated_at": data["generated_at"], "tmall_latest": data["months"][-1], "jd_latest": data["jd"]["latest_month"], "dy_latest": data["dy"]["months"][-1]}, ensure_ascii=False))


if __name__ == "__main__":
    main()

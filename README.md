# 市场洞察 BI 看板

GitHub Pages：<https://koralgrey.github.io/market-bi-dashboard/>

## 更新数据

支持单独更新一个平台；未提供新源的平台保留原数据。例如仅更新天猫类目销售额：

```bash
python scripts/update_dashboard.py \
  --tmall "/path/to/天猫_三级类目月度销售额.xlsx" \
  --rules-db "/path/to/运营资料库.sqlite"
```

`--rules-db`只读查询当前BSM/MSR版本，不修改资料库。类目汇总表没有品牌、店铺及商品字段，不能用它更新品牌市占或验证店铺身份。页面分别显示类目与品牌样本的截止月份，历史品牌样本保留但标注待当前规则复核。`--audit-report "/path/to/audit.json"`可另存逐月汇总校验与历史修订审计；不要将本地数据库或源Excel提交到公开仓库。

2026-10-27更新：天猫基础建材315类、家装主材268类，共583类，覆盖2025-01—2026-09，金额单位为元；源“宠物”页不在本建材看板范围。两个分表全部21个月已逐月与源汇总页回勾。京东、抖音保留2026-07数据；天猫品牌历史样本截至2026-06，尚未按BSM-1.14重算。

1. 安装 Python 3，然后执行 `pip install -r requirements.txt`。
2. 准备天猫、京东、抖音三份 Excel 源数据。
3. 在仓库目录运行：

```bash
python scripts/update_dashboard.py \
  --tmall "/path/to/天猫数据.xlsx" \
  --jd "/path/to/京东数据.xlsx" \
  --dy "/path/to/抖音数据.xlsx"
```

4. 本地打开 `index.html` 检查结果，然后提交并推送：

```bash
git add index.html
git commit -m "更新市场数据"
git push
```

推送到 `main` 后，GitHub Pages 会自动更新。

## 固定口径

- 天猫：基础建材、家装主材三级类目月度数据。
- 京东：只取月度、排除刷单；看板支持全部、自营、POP切换。
- 京东“乳胶漆”默认映射为“内墙乳胶漆”。
- 抖音防水喷剂与防水涂料分开；艺术漆合并；不同路径的置物架不合并。
- 抖音源表未提供品牌明细时，品牌分析明确显示“未提供品牌明细”。
- Excel源文件不提交到公开仓库，只提交聚合后的看板。

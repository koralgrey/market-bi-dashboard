# 市场洞察 BI 看板

GitHub Pages：<https://koralgrey.github.io/market-bi-dashboard/>

## 更新数据

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

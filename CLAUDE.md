# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

顺丰快递月度费用核对 Web 应用。接受 8 个 Excel 文件（账单/订单/城市/价格/B站赠送/KOC/出入库/售后），自动匹配店铺、省份，计算运费，与账单应付金额核对，输出结果。

## 启动命令

```bash
cd sf-freight-check
python app.py
# 访问 http://127.0.0.1:5000
```

## 依赖

- Python 3.x
- Flask
- openpyxl

## 核心架构

### 8个上传文件的角色

| 文件 | 作用 |
|------|------|
| 顺丰月结账单 | 主数据源，含运单号码/到件地区/产品类型/计费重量/应付金额 |
| 订单明细表 | 快递单号→店铺名称 |
| 全国城市省份表 | 城市→省份映射 |
| 顺丰物流价格表 | 按省份+重量段计算运费 |
| B站赠送明细表 | 发出快递单号→型号（B站抽奖渠道） |
| 推广KOC明细表 | 发出/寄回快递单号→型号（KOC渠道） |
| 出入库单明细表 | 出入库单号→备注（[其他]吉他情报局渠道） |
| 售后单明细表 | 退回快递单号→店铺名称 |

### 核心处理流程（cleaning.py）

1. **到件地区清洗** `normalize_area()` — 按 `/、,，-` 拆分多城市地区
2. **级联店铺匹配** `cascade_shop_match()` — 依次尝试：订单明细→售后→KOC→B站
3. **省份推断** `guess_province()` — 优先查映射表，兜底用 `PROVINCE_HINTS`
4. **运费计算** `calc_freight_from_table()` — 首重+续重，转寄/退回打6折
5. **账单核对** `process_bill()` — 对每行追加 店铺/省份/运费/总运费/是否一致/是否异常/备注

### 输出新列

`店铺 / 省份 / 运费 / 上浮费 / 总运费 / 是否一致？ / 是否异常？ / 备注`

## 关键文件

- `app.py` — Flask 路由，上传/处理/下载会话管理
- `utils/cleaning.py` — 所有数据清洗和业务逻辑
- `templates/index.html` — 前端页面（jinja2 模板）
- `exports/` — 临时导出文件目录（已 gitignore）

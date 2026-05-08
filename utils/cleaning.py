# -*- coding: utf-8 -*-
"""数据清洗和匹配逻辑"""

import re
from openpyxl import load_workbook


# ============================================================
# 城市→省份兜底表
# ============================================================
PROVINCE_HINTS = {
    '北京': '北京', '天津': '天津', '上海': '上海', '重庆': '重庆',
    '广州': '广东', '深圳': '广东', '东莞': '广东', '佛山': '广东',
    '惠州': '广东', '中山': '广东', '珠海': '广东', '江门': '广东',
    '成都': '四川', '绵阳': '四川', '德阳': '四川', '眉山': '四川',
    '资阳': '四川', '乐山': '四川', '泸州': '四川', '南充': '四川',
    '西安': '陕西', '咸阳': '陕西', '宝鸡': '陕西', '延安': '陕西',
    '武汉': '湖北', '襄阳': '湖北', '宜昌': '湖北', '荆州': '湖北',
    '长沙': '湖南', '岳阳': '湖南', '株洲': '湖南',
    '南京': '江苏', '苏州': '江苏', '无锡': '江苏', '常州': '江苏',
    '南通': '江苏', '徐州': '江苏', '扬州': '江苏',
    '杭州': '浙江', '宁波': '浙江', '温州': '浙江', '嘉兴': '浙江',
    '金华': '浙江', '台州': '浙江', '绍兴': '浙江',
    '郑州': '河南', '洛阳': '河南', '新乡': '河南', '南阳': '河南',
    '济南': '山东', '青岛': '山东', '烟台': '山东', '潍坊': '山东',
    '临沂': '山东', '淄博': '山东',
    '沈阳': '辽宁', '大连': '辽宁', '鞍山': '辽宁', '锦州': '辽宁',
    '哈尔滨': '黑龙江', '大庆': '黑龙江', '齐齐哈尔': '黑龙江',
    '长春': '吉林', '吉林': '吉林',
    '合肥': '安徽', '芜湖': '安徽', '蚌埠': '安徽',
    '南昌': '江西', '赣州': '江西', '九江': '江西',
    '福州': '福建', '厦门': '福建', '泉州': '福建', '漳州': '福建',
    '昆明': '云南', '贵阳': '贵州', '遵义': '贵州',
    '南宁': '广西', '桂林': '广西', '柳州': '广西',
    '海口': '海南', '三亚': '海南',
    '石家庄': '河北', '保定': '河北', '唐山': '河北', '廊坊': '河北',
    '太原': '山西', '大同': '山西',
    '呼和浩特': '内蒙古', '包头': '内蒙古',
    '兰州': '甘肃', '天水': '甘肃',
    '乌鲁木齐': '新疆', '克拉玛依': '新疆',
    '银川': '宁夏', '石嘴山': '宁夏',
    '西宁': '青海', '格尔木': '青海',
    '拉萨': '西藏',
}


# ============================================================
# 到件地区清洗
# ============================================================
def normalize_area(area):
    """
    拆分到件地区。
    输入: '成都/资阳/眉山' 或 '东莞'
    返回: (原始值, 主城市, 城市列表)
    """
    if not area:
        return '', '', []
    text = str(area).strip()
    parts = [p.strip() for p in re.split(r'[/、,，\-]', text) if p.strip()]
    main = parts[0] if parts else text
    return text, main, parts


# ============================================================
# 加载城市→省份映射
# ============================================================
def load_city_province_map(path):
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            mapping[str(row[0]).strip()] = str(row[1]).strip()
    return mapping


def guess_province(city, city_map):
    if not city:
        return ''
    city = str(city).strip()
    if city in city_map:
        return city_map[city]
    for k, v in PROVINCE_HINTS.items():
        if city.startswith(k):
            return v
    return ''


# ============================================================
# 加载订单明细（快递单号 → 店铺名称）
# ============================================================
def load_order_shop_map(path):
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {name: i for i, name in enumerate(header)}

    shop_col   = idx.get('店铺名称', -1)
    express_col = idx.get('快递单号', -1)

    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        if express_col != -1 and express_col < len(row) and row[express_col]:
            key = str(row[express_col]).strip()
            shop = str(row[shop_col]).strip() if (shop_col != -1 and shop_col < len(row) and row[shop_col]) else ''
            result[key] = shop
    return result


# ============================================================
# 加载价格表
# ============================================================
def load_price_table(path):
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {name: i for i, name in enumerate(header)}
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        item = {col: (row[i] if i < len(row) else None) for col, i in idx.items()}
        rows.append(item)
    return rows


def calc_freight_from_table(price_rows, product, service, province, weight):
    """
    运费公式（按价格表）：
      运费 = 首重价格 + (计件重量 - 首重重量) × 续重倍率
      转寄/退回 = 运费 × 0.6
    返回 (freight, surcharge, matched)
    """
    if weight is None:
        return None, None, False
    try:
        g = float(weight)
    except (TypeError, ValueError):
        return None, None, False

    for row in price_rows:
        if (str(row.get('产品类型', '')).strip() != str(product).strip()
                or str(row.get('到达省份', '')).strip() != str(province).strip()):
            continue
        try:
            first_w = float(row.get('首重重量（kg）') or 0)
            first_p = float(row.get('首重价格（元）') or 0)
            step_raw = row.get('续重倍率')
            step_p = float(step_raw) if step_raw is not None and str(step_raw).strip() not in ('', '/') else 0
        except (TypeError, ValueError):
            continue

        if g <= first_w:
            base = first_p
        else:
            base = first_p + (g - first_w) * step_p

        if str(service).strip() == '转寄/退回':
            freight = base * 0.6
        else:
            freight = base

        return round(freight, 2), None, True

    return None, None, False


# ============================================================
# 加载4张新匹配表：B站 / KOC / 出入库 / 售后
# ============================================================
def load_bill_matching_maps(bizhan_path, koc_path, io_path, aftersale_path):
    """
    加载所有新匹配表，返回字典：
    {
        'bizhan':   {快递单号 -> 实际发出型号},   # B站赠送
        'koc':      {快递单号 -> KOC型号},          # 推广KOC（发出+寄回）
        'io':       {出入库单号 -> 备注},           # 出入库明细
        'aftersale':{退回快递单号 -> 店铺名称},     # 售后明细
    }
    """
    return {
        'bizhan':    load_bizhan_map(bizhan_path),
        'koc':       load_koc_map(koc_path),
        'io':        load_io_map(io_path),
        'aftersale': load_aftersale_map(aftersale_path),
    }


def _header_idx(ws, keywords, row_start=1, row_end=5):
    """在指定行范围查找含有关键字的列索引"""
    for r in range(row_start, row_end + 1):
        row = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=r, max_row=r, values_only=True))]
        if all(any(k in v for v in row) for k in keywords):
            return {name: i for i, name in enumerate(row) if name}
    return None


def load_bizhan_map(path):
    """
    B站赠送明细表（其它额外建单发货协同表）：
    - Row1 是标题描述行，Row2 是真正的列头
    - 关键列：发出快递单号（index 9）、实际发出型号（index 12）
    - 匹配到 → 店铺="B站抽奖"，备注追加型号
    """
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 找到真正的列头行（Row2 含有"日期"、"要发的型号"等）
    header = None
    header_row = 1
    for i in range(1, 6):
        row = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=i, max_row=i, values_only=True))]
        if any('要发的型号' in v or '发出快递单号' in v for v in row):
            header = row
            header_row = i
            break
    if not header:
        return {}

    idx = {name: i for i, name in enumerate(header)}
    sf_col  = idx.get('发出快递单号', -1)
    model_col = idx.get('实际发出型号', -1)

    result = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or all(v is None for v in row):
            continue
        if sf_col != -1 and sf_col < len(row) and row[sf_col]:
            key = str(row[sf_col]).strip()
            model = str(row[model_col]).strip() if (model_col != -1 and model_col < len(row) and row[model_col]) else ''
            result[key] = model
    return result


def load_koc_map(path):
    """
    推广KOC明细表（推广协同表）：
    - Row1 是标题行，Row2 是列头（含"要发的型号"、"发出快递单号"等）
    - 关键列：发出快递单号（index 8）、寄回快递单号（index 11）、要发的型号（index 1）
    - 匹配到 → 店铺="推广KOC"，备注追加型号
    """
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    header = None
    header_row = 1
    for i in range(1, 6):
        row = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=i, max_row=i, values_only=True))]
        if any('要发的型号' in v or '发出快递单号' in v for v in row):
            header = row
            header_row = i
            break
    if not header:
        return {}

    idx = {name: i for i, name in enumerate(header)}
    send_col    = idx.get('发出快递单号', -1)
    return_col   = idx.get('寄回快递单号', -1)
    model_col    = idx.get('要发的型号', -1)

    result = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or all(v is None for v in row):
            continue
        model = str(row[model_col]).strip() if (model_col != -1 and model_col < len(row) and row[model_col]) else ''
        for col in [send_col, return_col]:
            if col != -1 and col < len(row) and row[col]:
                key = str(row[col]).strip()
                if key:
                    result[key] = model
    return result


def load_io_map(path):
    """
    出入库明细表：
    - 关键列：出入库单号（index 0）、备注（index 17）
    - 用于：当店铺=[其他]吉他情报局 时，匹配出入库单号并取备注
    """
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 找列头（第1行是标准列头）
    header = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
    idx = {name: i for i, name in enumerate(header)}

    io_col  = idx.get('出入库单号', -1)
    note_col = idx.get('备注', -1)

    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        if io_col != -1 and io_col < len(row) and row[io_col]:
            key = str(row[io_col]).strip()
            note = str(row[note_col]).strip() if (note_col != -1 and note_col < len(row) and row[note_col]) else ''
            result[key] = note
    return result


def load_aftersale_map(path):
    """
    售后明细表：
    - 关键列：退回快递单号（index 47）、店铺名称（index 0）
    - 匹配到 → 返回店铺名称
    """
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    # 找列头（第1行是标准列头）
    header = None
    header_row = 1
    for i in range(1, 6):
        row = [str(v).strip() if v is not None else '' for v in next(ws.iter_rows(min_row=i, max_row=i, values_only=True))]
        if any('退回快递单号' in v or '退货单号' in v for v in row):
            header = row
            header_row = i
            break
    if not header:
        return {}

    idx = {name: i for i, name in enumerate(header)}
    sf_col  = idx.get('退回快递单号', -1)
    shop_col = idx.get('店铺名称', -1)

    result = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if not row or all(v is None for v in row):
            continue
        if sf_col != -1 and sf_col < len(row) and row[sf_col]:
            key = str(row[sf_col]).strip()
            shop = str(row[shop_col]).strip() if (shop_col != -1 and shop_col < len(row) and row[shop_col]) else ''
            result[key] = shop
    return result


# ============================================================
# 级联店铺匹配
# ============================================================
def cascade_shop_match(waybill_no, order_map, match_maps):
    """
    店铺级联匹配：
    1. 订单明细表 → 快递单号匹配 → 店铺名称
    2. 售后明细表 → 退回快递单号匹配 → 店铺名称
    3. 推广KOC明细 → 发出/寄回快递单号匹配 → "推广KOC" + 备注追加型号
    4. B站赠送明细 → 发出快递单号匹配 → "B站抽奖" + 备注追加型号

    返回 (shop, extra_note, match_level)
      shop: 匹配到的店铺名（""表示未匹配）
      extra_note: 需要追加到备注的内容
      match_level: 匹配阶段（1~4）
    """
    if not waybill_no:
        return '', '', 0

    wb_key = str(waybill_no).strip()

    # 1. 订单明细
    if wb_key in order_map:
        shop = order_map[wb_key]
        if shop:
            return shop, '', 1

    # 2. 售后明细（退回快递单号）
    aftersale = match_maps.get('aftersale', {})
    if wb_key in aftersale:
        shop = aftersale[wb_key]
        if shop:
            return shop, '', 2

    # 3. 推广KOC（发出/寄回快递单号）
    koc = match_maps.get('koc', {})
    if wb_key in koc:
        model = koc[wb_key]
        return '推广KOC', f'要发的型号:{model}', 3

    # 4. B站赠送（发出快递单号）
    bizhan = match_maps.get('bizhan', {})
    if wb_key in bizhan:
        model = bizhan[wb_key]
        return 'B站抽奖', f'实际发出型号:{model}', 4

    return '', '', 0


# ============================================================
# 出入库单号&备注匹配
# ============================================================
def match_io_note(shop, waybill_no, order_map, match_maps, existing_remark):
    """
    仅当 店铺=[其他]吉他情报局 时执行：
    查找订单明细表获取 出入库单号，再查出入库明细表获取备注。
    返回追加到备注的内容。
    """
    if not waybill_no:
        return ''

    if shop and '[其他]' in str(shop) and '吉他情报局' in str(shop):
        # 直接从账单传入的shop已经是[其他]吉他情报局的情况
        pass
    else:
        # 先查订单明细获取店铺
        wb_key = str(waybill_no).strip()
        resolved_shop = order_map.get(wb_key, '')
        if not (resolved_shop and '[其他]' in str(resolved_shop) and '吉他情报局' in str(resolved_shop)):
            return ''

    # 查出入库明细
    io_map = match_maps.get('io', {})
    io_note = io_map.get(str(waybill_no).strip(), '')
    return f'出入库单号:{waybill_no}; {io_note}' if io_note else f'出入库单号:{waybill_no}'


# ============================================================
# 主处理函数
# ============================================================
def process_bill(bill_path, order_map, city_map, price_rows, match_maps):
    """
    处理账单明细，返回 (workbook, 处理结果列表, 统计摘要)

    新增列：
      店铺 / 省份 / 运费 / 上浮费 / 总运费 / 是否一致？ / 是否异常？ / 备注
    """
    wb = load_workbook(bill_path)
    ws = wb['账单明细']

    header = [cell.value for cell in ws[2]]
    col_idx = {str(v).strip(): i + 1 for i, v in enumerate(header) if v is not None}

    new_cols = ['店铺', '省份', '运费', '上浮费', '总运费', '是否一致？', '是否异常？', '备注']
    start_col = ws.max_column + 1
    for i, col in enumerate(new_cols, start=start_col):
        ws.cell(row=2, column=i, value=col)

    processed = []
    stats = {
        'total': 0, 'abnormal': 0,
        'consistent_ok': 0, 'consistent_fail': 0,
        'matched_shop': 0, 'matched_province': 0, 'matched_price': 0,
        'shop_order': 0, 'shop_aftersale': 0, 'shop_koc': 0, 'shop_bizhan': 0,
    }

    for r in range(3, ws.max_row + 1):
        row_empty = all(ws.cell(r, c).value is None for c in range(1, ws.max_column + 1))
        if row_empty:
            continue

        stats['total'] += 1
        remarks = []

        waybill_no = ws.cell(r, col_idx.get('运单号码', 3)).value
        area       = ws.cell(r, col_idx.get('到件地区', 5)).value
        product    = ws.cell(r, col_idx.get('产品类型', 8)).value
        service    = ws.cell(r, col_idx.get('服务', 14)).value
        weight     = ws.cell(r, col_idx.get('计费重量', 7)).value
        total_orig = ws.cell(r, col_idx.get('应付金额', 12)).value

        # 1. 到件地区清洗
        raw_area, main_city, city_list = normalize_area(area)

        # 2. 级联店铺匹配
        shop, extra_note, match_level = cascade_shop_match(waybill_no, order_map, match_maps)
        if match_level == 1:
            stats['shop_order'] += 1
        elif match_level == 2:
            stats['shop_aftersale'] += 1
        elif match_level == 3:
            stats['shop_koc'] += 1
        elif match_level == 4:
            stats['shop_bizhan'] += 1

        # 3. 出入库单号匹配（仅当店铺含[其他]吉他情报局）
        io_extra = match_io_note(shop, waybill_no, order_map, match_maps, '')
        if io_extra:
            remarks.append(io_extra)

        # 4. 追加KOC/B站型号备注
        if extra_note:
            remarks.append(extra_note)

        # 5. 省份匹配
        province = guess_province(main_city, city_map)

        # 6. 运费计算
        freight, _, price_matched = calc_freight_from_table(
            price_rows,
            str(product).strip() if product else '',
            str(service).strip() if service else '运费',
            province,
            weight,
        )

        surcharge = None   # 上浮费留空

        shop_ok    = bool(shop)
        province_ok = bool(province)
        price_ok   = freight is not None

        if shop_ok:
            stats['matched_shop'] += 1
        else:
            remarks.append('未匹配店铺')
        if province_ok:
            stats['matched_province'] += 1
        else:
            remarks.append('未匹配省份')
        if price_ok:
            stats['matched_price'] += 1
        else:
            remarks.append('未匹配价格')

        abnormal = '异常' if (not shop_ok or not province_ok or not price_ok) else '正常'
        if abnormal == '异常':
            stats['abnormal'] += 1

        computed_total = freight

        consistent = ''
        if freight is not None and total_orig is not None:
            try:
                if round(float(freight), 2) == round(float(total_orig), 2):
                    consistent = '一致'
                    stats['consistent_ok'] += 1
                else:
                    consistent = '不一致'
                    stats['consistent_fail'] += 1
                    remarks.append(f'运费({round(float(freight), 2)})≠应付({total_orig})')
            except (TypeError, ValueError):
                consistent = ''

        # 到件地区多城市备注
        if raw_area != main_city and len(city_list) > 1:
            remarks.insert(0, f'到件地区含多城市:{raw_area}')

        remark_text = '; '.join(remarks)
        values = [shop, province, freight, surcharge, computed_total, consistent, abnormal, remark_text]
        for i, val in enumerate(values, start=start_col):
            ws.cell(row=r, column=i, value=val)

        processed.append([
            waybill_no, raw_area, shop, province, freight, surcharge,
            computed_total, consistent, abnormal, remark_text
        ])

    return wb, processed, stats


def export_result(wb, output_path):
    wb.save(output_path)

# -*- coding: utf-8 -*-
"""
顺丰物流月度费用核对 - Flask 应用入口
"""

from flask import Flask, render_template, request, send_file, session
from utils.cleaning import (
    load_city_province_map,
    load_order_shop_map,
    load_price_table,
    load_bill_matching_maps,
    process_bill,
    export_result,
)
import tempfile
import os
import uuid
from pathlib import Path

app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024   # 80MB

EXPORT_DIR = Path(__file__).parent / 'exports'
EXPORT_DIR.mkdir(exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        bill_file     = request.files.get('bill_file')
        order_file   = request.files.get('order_file')
        city_file    = request.files.get('city_file')
        price_file   = request.files.get('price_file')
        bizhan_file  = request.files.get('bizhan_file')
        koc_file     = request.files.get('koc_file')
        io_file      = request.files.get('io_file')
        aftersale_file = request.files.get('aftersale_file')

        if not all([bill_file, order_file, city_file, price_file,
                     bizhan_file, koc_file, io_file, aftersale_file]):
            return '请上传全部8个文件', 400

        # 保存到临时目录
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            bill_path      = td / bill_file.filename
            order_path    = td / order_file.filename
            city_path     = td / city_file.filename
            price_path    = td / price_file.filename
            bizhan_path   = td / bizhan_file.filename
            koc_path      = td / koc_file.filename
            io_path       = td / io_file.filename
            aftersale_path = td / aftersale_file.filename

            bill_file.save(str(bill_path))
            order_file.save(str(order_path))
            city_file.save(str(city_path))
            price_file.save(str(price_path))
            bizhan_file.save(str(bizhan_path))
            koc_file.save(str(koc_path))
            io_file.save(str(io_path))
            aftersale_file.save(str(aftersale_path))

            # 加载所有参考数据
            city_map  = load_city_province_map(str(city_path))
            order_map = load_order_shop_map(str(order_path))
            price_rows = load_price_table(str(price_path))
            match_maps = load_bill_matching_maps(
                str(bizhan_path),
                str(koc_path),
                str(io_path),
                str(aftersale_path),
            )

            # 处理账单
            wb, processed, stats = process_bill(
                str(bill_path),
                order_map,
                city_map,
                price_rows,
                match_maps,
            )

        # 生成导出文件
        export_id   = uuid.uuid4().hex[:8]
        export_name = f'顺丰核对结果_{export_id}.xlsx'
        export_path = EXPORT_DIR / export_name
        export_result(wb, str(export_path))

        session['export_file'] = str(export_path)
        session['export_name'] = export_name

        preview_cols = ['运单号码', '到件地区', '店铺', '省份', '运费', '上浮费', '总运费', '是否一致？', '是否异常？', '备注']
        preview_rows = processed[:20]

        return render_template(
            'index.html',
            stats=stats,
            download_url='/download',
            preview_cols=preview_cols,
            preview_rows=preview_rows,
        )

    return render_template('index.html', stats=None)


@app.route('/download')
def download():
    export_path = session.get('export_file')
    export_name = session.get('export_name', '顺丰核对结果.xlsx')
    if not export_path or not Path(export_path).exists():
        return '文件不存在，请重新上传处理', 404
    return send_file(export_path, as_attachment=True, download_name=export_name)


@app.route('/clear')
def clear():
    session.pop('export_file', None)
    session.pop('export_name', None)
    return render_template('index.html', stats=None)


if __name__ == '__main__':
    app.run(debug=True, port=5000, host='0.0.0.0')

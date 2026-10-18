# -*- coding: utf-8 -*-
"""解析層測試：由真實 .xlsx 走完 parse_excel，確認欄位偵測與 Total 行排除。"""

import io

from openpyxl import Workbook

import app as A


def make_xlsx(rows, header):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Data'
    ws.append(list(header))
    for r in rows:
        ws.append(list(r))
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def test_parse_shop_level_file():
    data = make_xlsx([
        ('2025-01', 'ZZ99 - 測試店', 10, 80, 48, 32, 0.4),
        ('2026-01', 'ZZ99 - 測試店', 12, 100, 60, 40, 0.4),
        # Total 行必須被忽略
        ('Total', '', 22, 180, 108, 72, 0.4),
    ], ('Month', 'Shop', 'Sold Qty', 'Net Sales', 'COGS', 'GP$', 'GP%'))

    ds = A.parse_excel(data)

    assert ds['kind'] == 'shop'
    assert ds['months'] == ['2025-01', '2026-01']
    assert ds['siteCount'] == 1
    assert len(ds['rows']) == 2
    assert ds['sites'][0]['code'] == 'ZZ99'
    assert ds['sites'][0]['name'] == '測試店'

    # rows = [月份索引, 店鋪索引, 銷量, 淨銷售, 毛利]
    first = ds['rows'][0]
    assert first[0] == 0 and first[1] == 0
    assert first[2:] == [10, 80.0, 32.0]


def test_parse_gp_falls_back_to_sales_minus_cogs():
    data = make_xlsx([
        ('2026-01', 'ZZ98 - 無GP欄', 3, 100, 70),
    ], ('Month', 'Shop', 'Sold Qty', 'Net Sales', 'COGS'))

    ds = A.parse_excel(data)
    assert ds['rows'][0][4] == 30.0


def test_parse_rich_file_exposes_dimensions():
    data = make_xlsx([
        ('2026-01', 'ZZ99 - 測試店', 1, 100, 60, 40, 'BrandA', 'MC1'),
        ('2026-01', 'ZZ99 - 測試店', 2, 200, 120, 80, 'BrandB', 'MC2'),
    ], ('Month', 'Shop', 'Sold Qty', 'Net Sales', 'COGS', 'GP$', 'Brand', 'MC'))

    ds = A.parse_excel(data)

    assert ds['kind'] == 'rich'
    assert set(ds['dims']) >= {'brand', 'mc'}
    assert ds['dims']['brand']['values'] == ['BrandA', 'BrandB']
    # rich 檔的 rows 末端會追加各維度索引
    assert len(ds['rows'][0]) > 5


def test_parse_rejects_file_without_header():
    data = make_xlsx([('foo', 'bar')], ('A', 'B'))
    try:
        A.parse_excel(data)
    except ValueError as e:
        assert '找不到資料表' in str(e)
    else:
        raise AssertionError('應該要拋出 ValueError')

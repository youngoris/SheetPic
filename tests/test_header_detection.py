"""Tests for header-row detection (`_score_header_row`)."""
import os
import sys
import datetime as dt
import tempfile

import pandas as pd
import pytest

# Make sheetpic.py importable
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from sheetpic import _score_header_row, _cell_type  # noqa: E402


def df(rows):
    """Build a DataFrame (header=None style) from a list of row lists."""
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# _cell_type
# ---------------------------------------------------------------------------

def test_cell_type_basic():
    assert _cell_type(None) == 'blank'
    assert _cell_type('') == 'blank'
    assert _cell_type('   ') == 'blank'
    assert _cell_type(123) == 'num'
    assert _cell_type(1.5) == 'num'
    assert _cell_type('123') == 'num'  # numeric string
    assert _cell_type('1,234.5') == 'num'
    assert _cell_type('hello') == 'str'
    assert _cell_type('https://example.com/x.jpg') == 'url'
    assert _cell_type('//cdn.example.com/x.jpg') == 'url'
    assert _cell_type(dt.date(2024, 1, 1)) == 'date'
    assert _cell_type(dt.datetime(2024, 1, 1)) == 'date'


# ---------------------------------------------------------------------------
# _score_header_row — synthetic shapes
# ---------------------------------------------------------------------------

def test_simple_header_at_row0():
    data = df([
        ['SKU', 'Name', 'Price', 'Image'],
        ['A001', 'Widget', 9.9, 'https://x/1.jpg'],
        ['A002', 'Gadget', 12.5, 'https://x/2.jpg'],
        ['A003', 'Thing', 3.0, 'https://x/3.jpg'],
    ])
    assert _score_header_row(data) == 0


def test_header_after_title_row():
    data = df([
        ['Q1 Sales Report', None, None, None],
        ['SKU', '名称', '价格', '图片'],
        ['A001', '小部件', 9.9, 'https://x/1.jpg'],
        ['A002', '小工具', 12.5, 'https://x/2.jpg'],
        ['A003', '东西', 3.0, 'https://x/3.jpg'],
    ])
    assert _score_header_row(data) == 1


def test_header_after_multiple_title_rows():
    """Common pattern: title + subtitle + blank + header + data."""
    data = df([
        ['瑞美家供应商商品列表', None, None, None, None],
        ['导出时间: 2024-01-01', None, None, None, None],
        [None, None, None, None, None],
        ['货号', '品名', '规格', '单价', '主图'],
        ['SKU01', '杯子', '大', 19.9, 'https://i/1.jpg'],
        ['SKU02', '盘子', '中', 12.0, 'https://i/2.jpg'],
        ['SKU03', '碗',   '小',  8.5, 'https://i/3.jpg'],
        ['SKU04', '勺子', '小',  5.0, 'https://i/4.jpg'],
    ])
    assert _score_header_row(data) == 3


def test_header_at_row4_with_merged_title():
    """Mimics merged top cells (only first cell filled in title rows)."""
    data = df([
        ['公司年度报表', None, None, None, None, None],
        [None, None, None, None, None, None],
        ['部门: 销售', None, None, None, None, None],
        [None, None, None, None, None, None],
        ['编号', '商品名称', '颜色', '尺寸', '价格', '库存'],
        [1, '椅子', '红', 'L', 199, 50],
        [2, '桌子', '白', 'M', 299, 30],
        [3, '柜子', '黑', 'S', 399, 20],
        [4, '床',   '灰', 'XL', 999, 10],
    ])
    assert _score_header_row(data) == 4


def test_url_only_column_does_not_promote_data_row():
    """Data row that's mostly strings should NOT beat the real header."""
    data = df([
        ['Code', 'Name', 'Color', 'URL'],
        ['A1', 'Foo', 'red', 'https://x/a.jpg'],
        ['A2', 'Bar', 'blue', 'https://x/b.jpg'],
        ['A3', 'Baz', 'green', 'https://x/c.jpg'],
    ])
    assert _score_header_row(data) == 0


def test_pure_data_no_header_returns_first_filled():
    """Robustness: no obvious header → don't crash, return something sane (≤1)."""
    data = df([
        [1, 2.0, 'x'],
        [2, 3.0, 'y'],
        [3, 4.0, 'z'],
    ])
    # Without a real header, returning row 0 is acceptable.
    assert _score_header_row(data) in (0, 1, 2)


def test_empty_first_rows_skipped():
    data = df([
        [None, None, None, None],
        [None, None, None, None],
        ['SKU', 'Name', 'Qty', 'Price'],
        ['S1', 'A', 10, 1.0],
        ['S2', 'B', 20, 2.0],
        ['S3', 'C', 30, 3.0],
    ])
    assert _score_header_row(data) == 2


def test_long_description_title_not_chosen():
    """A long-text title row should lose to a short-label header row."""
    data = df([
        ['本表统计了2024年第一季度的全部销售数据，包括各分公司、各品类、各SKU的销售额', None, None, None],
        ['SKU', '名称', '数量', '金额'],
        ['A1', '产品一', 10, 100],
        ['A2', '产品二', 20, 200],
        ['A3', '产品三', 30, 300],
    ])
    assert _score_header_row(data) == 1


def test_header_with_chinese_keywords():
    data = df([
        ['报表', None, None, None, None],
        ['序号', '商品编码', '商品名称', '图片链接', '售价'],
        [1, 'SKU001', '杯子',  'https://i/1.jpg', 19.9],
        [2, 'SKU002', '盘子',  'https://i/2.jpg', 12.0],
        [3, 'SKU003', '碗',    'https://i/3.jpg',  8.5],
    ])
    assert _score_header_row(data) == 1


def test_two_rows_of_strings_picks_actual_header():
    """When title row + header row both look text-heavy, pick the LATER one
    (the one followed by data)."""
    data = df([
        ['Quarterly', 'Sales', 'Report', 'Confidential'],
        ['SKU', 'Name', 'Price', 'Image'],
        ['A1', 'Foo', 1.0, 'https://x/1.jpg'],
        ['A2', 'Bar', 2.0, 'https://x/2.jpg'],
        ['A3', 'Baz', 3.0, 'https://x/3.jpg'],
    ])
    assert _score_header_row(data) == 1


def test_real_xlsx_roundtrip(tmp_path):
    """Write a real xlsx and verify detection through pandas.read_excel."""
    rows = [
        ['店铺销售明细', None, None, None, None],
        ['日期: 2024-05-01', None, None, None, None],
        [None, None, None, None, None],
        ['SKU', '名称', '颜色', '价格', '主图'],
        ['A001', '杯子', '红', 19.9, 'https://i/1.jpg'],
        ['A002', '盘子', '白', 12.0, 'https://i/2.jpg'],
        ['A003', '碗',   '黑',  8.5, 'https://i/3.jpg'],
    ]
    p = tmp_path / 'sample.xlsx'
    pd.DataFrame(rows).to_excel(p, header=False, index=False)

    raw = pd.read_excel(p, header=None, nrows=40)
    assert _score_header_row(raw) == 3


def test_single_row_returns_zero():
    data = df([['A', 'B', 'C']])
    # Single row has no "data follows" support; algo can return 0.
    assert _score_header_row(data) == 0


def test_empty_df_returns_zero():
    assert _score_header_row(pd.DataFrame()) == 0


def test_numeric_first_row_loses_to_text_header():
    """If the first row is numeric (e.g., column index numbers), the actual
    header below it should be selected."""
    data = df([
        [1, 2, 3, 4, 5],
        ['SKU', 'Name', 'Color', 'Size', 'Price'],
        ['A1', 'Foo', 'red', 'L', 9.9],
        ['A2', 'Bar', 'blue', 'M', 12.0],
        ['A3', 'Baz', 'green', 'S', 5.0],
    ])
    assert _score_header_row(data) == 1


def test_duplicate_string_row_not_preferred():
    """Repeated values in a string-heavy row are unlikely to be a header."""
    data = df([
        ['同上', '同上', '同上', '同上'],
        ['SKU', 'Name', 'Qty', 'Price'],
        ['A1', 'X', 1, 1.0],
        ['A2', 'Y', 2, 2.0],
        ['A3', 'Z', 3, 3.0],
    ])
    assert _score_header_row(data) == 1

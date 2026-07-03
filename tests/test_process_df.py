"""Tests for column-collision handling in process_df.

Reproduces a real-world failure: a sheet with several blank/duplicate header
cells (which pandas labels Unnamed: N) and duplicate names (e.g. two "条码"
columns) caused `df[name]` to return a DataFrame, breaking image-column
auto-detection.
"""
import os
import sys
import pandas as pd
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)


def _build_app(df, wb=None, ws=None):
    """Build a SheetPicApp instance without launching Tk, just enough to call
    process_df()."""
    from sheetpic import SheetPicApp, LANG_MAP
    app = SheetPicApp.__new__(SheetPicApp)
    app.T = LANG_MAP['zh']
    app.df = df.copy()
    app.wb = wb
    app.ws = ws
    app.mode = 'extract'
    app.sorted_img_cols = []
    app.embed_url_cols = []
    app.embed_url_col_idx = 0
    app.embed_sku_col_idx = 0
    app.embed_use_url_library = False
    app.url_library = {}
    app._url_library_combo_value = None

    # Stub root so process_df's after-call works for update_ui_lists.
    captured = {}

    class _RootStub:
        def after(self, _delay, fn, *a):
            try:
                fn(*a)
            except Exception:
                pass

    app.root = _RootStub()

    # Stub combos / log so update_ui_lists works.
    class _ComboStub:
        def __init__(self):
            self.values = []
            self.current_idx = -1
            self.value = None

        def __setitem__(self, k, v):
            if k == 'values':
                self.values = v

        def current(self, i):
            self.current_idx = i
            if 0 <= i < len(self.values):
                self.value = self.values[i]

        def set(self, v):
            self.value = v

        def config(self, **kw):
            pass

        def bind(self, *_args, **_kwargs):
            pass

    app.combo_img = _ComboStub()
    app.combo_code = _ComboStub()
    app.combo_url = _ComboStub()
    app.combo_sku = _ComboStub()

    class _ButtonStub:
        def __init__(self):
            self.state = None

        def config(self, **kw):
            self.state = kw.get('state', self.state)

    app.btn_run = _ButtonStub()
    app.log = lambda *_a, **_k: None
    app._captured = captured
    return app


def test_duplicate_unnamed_columns_does_not_crash():
    """4 unnamed + 2 same-name columns must not break URL detection."""
    df = pd.DataFrame({
        '链接': ['http://x/1.jpg', 'http://x/2.jpg', 'http://x/3.jpg'],
        '条码': ['111', '222', '333'],
        '商品名称': ['A', 'B', 'C'],
        '条码.1': ['111', '222', '333'],  # pandas-style duplicate
        'Unnamed: 4': [None, None, None],
        'Unnamed: 5': [None, None, None],
        'Unnamed: 6': [None, None, None],
        'Unnamed: 7': [None, None, None],
    })
    app = _build_app(df)
    app.process_df()
    # The URL column should have been detected as the first column.
    assert len(app.sorted_img_cols) == 1
    assert app.sorted_img_cols[0]['idx'] == 0
    assert app.sorted_img_cols[0]['count'] == 3
    assert app.sorted_img_cols[0]['type'] == 'url'


def test_columns_are_made_unique():
    df = pd.DataFrame({
        'A': [1, 2],
        'A.1': [3, 4],
        'Unnamed: 2': [5, 6],
        'Unnamed: 3': [7, 8],
    })
    app = _build_app(df)
    app.process_df()
    assert len(set(app.df.columns)) == len(app.df.columns)


def test_process_df_does_not_coerce_all_values_to_strings():
    df = pd.DataFrame({
        '链接': ['http://x/1.jpg', None],
        '数量': [1, 2],
    })
    app = _build_app(df)
    app.process_df()
    assert app.df.iloc[0, 1] == 1


def test_url_library_enables_embed_without_url_column():
    df = pd.DataFrame({
        '条码': ['A001', 'A002'],
        '商品名称': ['杯子', '盘子'],
    })
    app = _build_app(df)
    app.mode = 'embed'
    app.url_library = {'A001': 'https://img.example.com/a001.webp'}

    app.process_df()

    assert app.combo_url.values == ['[URL库] 按SKU/ID匹配 (1 条)']
    assert app.embed_use_url_library is True
    assert app.embed_url_col_idx is None
    assert app.btn_run.state == 'normal'


def test_url_library_import_indexes_all_identifier_columns():
    df = pd.DataFrame({
        'SKU': ['S001', 'S002'],
        '条形码': ['690001', '690002'],
        '商品名称': ['杯子', '盘子'],
        '图片': ['https://img.example.com/1.jpg', 'https://img.example.com/2.jpg'],
    })
    app = _build_app(df)

    added = app._merge_url_library_from_df(df)

    assert added == 4
    assert app.url_library['S001'] == 'https://img.example.com/1.jpg'
    assert app.url_library['690001'] == 'https://img.example.com/1.jpg'
    assert app.url_library['S002'] == 'https://img.example.com/2.jpg'
    assert app.url_library['690002'] == 'https://img.example.com/2.jpg'


def test_url_library_import_keeps_extra_fields_for_ui_and_output():
    df = pd.DataFrame({
        'SKU': ['S001', 'S002'],
        '条形码': ['690001', '690002'],
        '商品名称': ['杯子', '盘子'],
        '品牌': ['方寸', '方寸'],
        '图片': ['https://img.example.com/1.jpg', 'https://img.example.com/2.jpg'],
    })
    app = _build_app(df)

    added = app._merge_url_library_from_df(df)

    assert added == 4
    assert app.url_library_field_names == ['SKU', '条形码', '商品名称', '品牌']
    assert app.url_library_selected_fields == ['商品名称', '品牌']
    assert app.url_library_records['S001']['商品名称'] == '杯子'
    assert app.url_library_records['690002']['品牌'] == '方寸'


def test_url_library_auto_selects_matching_barcode_column():
    df = pd.DataFrame({
        '组合': ['货架组合1', None, None],
        '货位编码': ['H01-01-01-01', 'H01-01-01-A', 'H01-01-01-02'],
        '条形码': ['690001', None, '690002'],
        '商品名称': ['商品1', '分类行', '商品2'],
    })
    app = _build_app(df)
    app.mode = 'embed'
    app.url_library = {
        '690001': 'https://img.example.com/1.jpg',
        '690002': 'https://img.example.com/2.jpg',
    }

    app.process_df()

    assert app.combo_url.values == ['[URL库] 按SKU/ID匹配 (2 条)']
    assert app.embed_use_url_library is True
    assert app.combo_sku.value == '条形码 (C)'
    assert app.embed_sku_col_idx == 2


def test_embed_sku_default_prefers_barcode_and_syncs_index():
    df = pd.DataFrame({
        '组合': ['货架组合1', None],
        '货位编码': ['H01-01-01-01', 'H01-01-01-02'],
        '条形码': ['690001', '690002'],
        '图片': ['https://img.example.com/1.jpg', 'https://img.example.com/2.jpg'],
    })
    app = _build_app(df)
    app.mode = 'embed'

    app.process_df()

    assert app.combo_sku.value == '条形码 (C)'
    assert app.embed_sku_col_idx == 2


def test_real_user_file_qinrun(tmp_path):
    """The actual failing file shape: '链接' header + 4 trailing blank columns."""
    rows = [
        ['链接', '条码', '商品名称', '条码', '规格', '进价', '售价',
         None, '成本', None, None, None],
        ['http://erp/1.jpg!200x200', 6973854641429, '厨美利菠萝刀',
         6973854641429, '24/240', 6.6, 10.9, 6.6, 5, 0.24, None, None],
        ['http://erp/2.jpg!200x200', 6973854640613, '厨美利筛',
         6973854640613, '6/48', 13.8, 19.9, 13.8, 10.2, 0.26, None, None],
        ['http://erp/3.jpg!200x200', 6973854640637, '厨美利饭勺',
         6973854640637, '12/144', 8.5, 13.9, 8.5, 6.6, 0.22, None, None],
    ]
    p = tmp_path / 'qr.xlsx'
    pd.DataFrame(rows).to_excel(p, header=False, index=False)

    # Mirror the analyze_data flow.
    from sheetpic import _score_header_row
    raw = pd.read_excel(p, header=None, nrows=40)
    hr = _score_header_row(raw)
    assert hr == 0
    df = pd.read_excel(p, header=hr)

    app = _build_app(df)
    app.process_df()
    assert app.sorted_img_cols, 'image column must be detected'
    assert app.sorted_img_cols[0]['idx'] == 0
    assert app.sorted_img_cols[0]['count'] == 3

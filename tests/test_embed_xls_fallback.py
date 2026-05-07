import os
import sys
from io import BytesIO

import openpyxl
import pandas as pd
from PIL import Image as PILImage

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)


def test_write_original_xls_falls_back_to_new_xlsx(tmp_path):
    from sheetpic import LANG_MAP, SheetPicApp

    app = SheetPicApp.__new__(SheetPicApp)
    app.T = LANG_MAP['zh']
    app.file_path = str(tmp_path / 'legacy.xls')
    app.header_row = 2
    app.df = pd.DataFrame({
        '图片': ['http://x/1.jpg'],
        '条码': ['A001'],
    })
    logs = []

    class _RootStub:
        def after(self, _delay, fn, *args):
            fn(*args)

    class _EntryStub:
        def get(self):
            return str(tmp_path)

    with open(app.file_path, 'wb') as f:
        f.write(b'not an xlsx file')

    app.root = _RootStub()
    app.entry_dest = _EntryStub()
    app.log = logs.append

    out_file, ws, wb_out, img_header_col, header_row_excel = app._embed_setup_original(
        'legacy', 0, del_url=True
    )

    try:
        assert out_file == str(tmp_path / 'legacy_Embedded.xlsx')
        assert img_header_col == 1
        assert header_row_excel == 3
        assert ws.cell(row=3, column=1).value == '图片'
        assert app._embed_setup_used_original is False
        assert any('.xls' in msg for msg in logs)
    finally:
        wb_out.close()


def test_xls_fallback_run_writes_source_rows(tmp_path, monkeypatch):
    import sheetpic
    from sheetpic import LANG_MAP, SheetPicApp

    app = SheetPicApp.__new__(SheetPicApp)
    app.T = LANG_MAP['zh']
    app.file_path = str(tmp_path / 'legacy.xls')
    app.header_row = 0
    app.df = pd.DataFrame({
        '图片': ['http://x/1.jpg'],
        '条码': ['A001'],
        '商品全名': ['测试商品'],
    })
    app.embed_url_col_idx = 0
    app.embed_sku_col_idx = 1
    app.is_running = True
    app._process_start_time = 0
    logs = []
    status = []

    class _RootStub:
        def after(self, _delay, fn, *args):
            fn(*args)

    class _EntryStub:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    class _VarStub:
        def __init__(self, value):
            self.value = value

        def get(self):
            return self.value

    class _WidgetStub:
        def config(self, **kwargs):
            status.append(kwargs)

    class _ProgressStub(dict):
        def stop(self):
            pass

    with open(app.file_path, 'wb') as f:
        f.write(b'not an xlsx file')

    img_bytes = BytesIO()
    PILImage.new('RGB', (10, 10), 'white').save(img_bytes, format='PNG')
    img_payload = img_bytes.getvalue()

    def _download(_url, _max_dim=None):
        return True, BytesIO(img_payload)

    app.root = _RootStub()
    app.entry_dest = _EntryStub(str(tmp_path))
    app.entry_max_dim = _EntryStub('500')
    app.var_original = _VarStub(False)
    app.var_write_original = _VarStub(True)
    app.var_del_url = _VarStub(True)
    app.progress = _ProgressStub()
    app.lbl_status = _WidgetStub()
    app.btn_run = _WidgetStub()
    app.btn_stop = _WidgetStub()
    app.log = logs.append
    app.download_to_bytesio = _download
    app._open_folder = lambda _path: None
    monkeypatch.setattr(sheetpic.messagebox, 'showinfo', lambda *_args, **_kwargs: None)

    app.run_embed_process()

    out_file = tmp_path / 'legacy_Embedded.xlsx'
    assert out_file.exists()

    wb = openpyxl.load_workbook(out_file)
    try:
        ws = wb.active
        assert ws.cell(row=1, column=1).value == '图片'
        assert ws.cell(row=1, column=2).value == '条码'
        assert ws.cell(row=2, column=2).value == 'A001'
        assert ws.cell(row=2, column=3).value == '测试商品'
    finally:
        wb.close()

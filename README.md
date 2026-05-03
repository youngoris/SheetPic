# SheetPic v1.0

**Batch Image Extract & Embed Tool for Spreadsheets**
**表格图片批量提取 & 嵌入工具**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Built%20with-Python%203.10%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-macOS%20%7C%20Windows-lightgrey.svg)]()

---

## Features

Two modes in one app, switch via tabs:

### Extract (提取图片)

Download or export images from spreadsheets into a local folder.

- **Dual-source parsing**: `Pandas` for text/URL columns, `OpenPyXL` for embedded cell images
- **Smart header detection**: Auto-locates the header row even if data starts at row 5
- **Multi-column merge**: When multiple columns contain images, auto-selects the richest column
- **Clipboard mode**: Copy a table from anywhere, paste and process

### Embed (嵌入图片)

Download images from URLs and embed them directly into Excel cells.

- **URL column detection**: Auto-detects columns containing image URLs
- **Thumbnail parameter stripping**: Removes CDN resize params (`!200x200`, `?width=300`, `?x-oss-process=...`) to download originals
- **Aspect ratio preservation**: Images scale to fit row height while keeping original proportions
- **Configurable size**: Set max dimension (default 500px), or insert original resolution
- **Delete URL column**: Option to remove the source URL column after embedding

### Shared

- **Bilingual UI**: Auto-detects system language (Chinese / English)
- **Stop button**: Gracefully halt any running task
- **Transparent logs**: Real-time status with clear error messages (`404`, `Timeout`, `Empty`)
- **Anti-blocking**: Realistic User-Agent headers

---

## Screenshots

<img width="499" height="607" alt="image" src="https://github.com/user-attachments/assets/5f64aa56-1e2b-4b26-a95e-d1370af364f6" />

---

## Download

Go to [Releases](../../releases) and download:

| Platform | File | Notes |
|---|---|---|
| macOS (Apple Silicon) | `SheetPic-macOS-ARM64.dmg` | Open DMG, drag SheetPic to Applications |
| Windows x64 | `SheetPic-Windows-x64.zip` | Unzip, run `SheetPic.exe` inside |

---

## Usage

### Extract mode (提取)

1. Select an Excel/CSV file, or click **Clipboard** to paste data
2. Choose the **Image Source** column (auto-detected)
3. Choose the **Filename** column (e.g., SKU/Code)
4. Click **Start** -- images are saved to a folder

### Embed mode (嵌入)

1. Select an Excel/CSV file containing image URLs
2. Choose the **URL column** and **SKU/ID column**
3. Set **max dimension** (default 500px) or check **Original Size**
4. Optionally check **Delete URL column after embedding**
5. Click **Start** -- a new Excel file is created with images embedded in cells

---

## Development

```bash
git clone https://github.com/youngoris/SheetPic.git
cd SheetPic

pip install pandas openpyxl xlrd lxml requests Pillow PyInstaller

python sheetpic.py
```

### Build locally

```bash
# macOS (produces dist/SheetPic.dmg)
python3 build.py

# Windows (produces dist/SheetPic.exe)
python build.py
```

### CI/CD

Pushing a `v*` tag triggers GitHub Actions to build macOS ARM64 and Windows x64 binaries, then auto-creates a Release.

```bash
git tag v1.0.0
git push origin v1.0.0
```

---

## Tech Stack

- **GUI**: Tkinter + ttk
- **Excel**: OpenPyXL (read/write/embed)
- **Data**: Pandas (CSV/HTML/Excel parsing)
- **Download**: requests + ThreadPoolExecutor (10 workers)
- **Image**: Pillow (resize/format conversion)
- **Packaging**: PyInstaller

---

## License

MIT

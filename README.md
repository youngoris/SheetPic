# SheetPic v5.0 🚀

**The Ultimate Batch Image Downloader for E-commerce & Operations.**
**专为电商运营打造的表格图片批量下载/提取神器。**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Built%20with-Python%203.10%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey.svg)]()

---

## 📖 Introduction (简介)

**SheetPic** solves the nightmare of downloading thousands of product images from messy distributor spreadsheets. Whether the images are **embedded** in the Excel cells or provided as **URLs**, SheetPic handles them all.

它解决了电商运营中最头疼的问题：从混乱的供应商表格中提取图片。无论图片是**直接嵌入在单元格里**的，还是**HTTP 链接**，SheetPic 都能智能识别并批量下载。

## ✨ Key Features (核心功能)

### 🧠 1. Dual-Core Engine (双核引擎)
* **Universal Parsing**: Uses `Pandas` for robust text/URL reading (supports `.xlsx`, `.xls`, `.csv`, `.html`).
* **Embedded Extraction**: Uses `OpenPyXL` to extract images pasted directly into cells.
* **Clipboard Mode**: File corrupted? Just copy the table and click **"Read Clipboard"**.

### ⚡ 2. Smart & HD (智能与高清)
* **HD Quality**: Automatically strips thumbnail parameters (e.g., `!200x200`, `?width=300`) to ensure you get the **original high-res image**.
* **Smart Header Seek**: Automatically detects the header row, even if the table starts at row 5.
* **Multi-Column Merge**: If multiple columns contain images, it prioritizes the column with the most data and auto-renames duplicates (e.g., `SKU-1.jpg`).

### 🛡️ 3. Robustness (鲁棒性设计)
* **Stop Button**: Gracefully stop the task anytime without crashing.
* **Transparent Logs**: Clearly distinguishes between `[404 Not Found]`, `[Timeout]`, and `[Empty]` cells.
* **Smart Resume**: Skips empty rows instantly to save time.
* **Anti-Blocking**: Uses realistic User-Agent headers to prevent server rejection.

---

## 📸 Screenshots (界面预览)

<img width="499" height="607" alt="image" src="https://github.com/user-attachments/assets/5f64aa56-1e2b-4b26-a95e-d1370af364f6" />


> **UI Philosophy**: Compact card-style layout with high-contrast buttons and a vivid green progress bar.

---

## 📥 Installation & Usage (安装与使用)

### For Users (直接使用)
1.  Go to [Releases](../../releases) and download `SheetPic_v5.exe`.
2.  Run the app (No installation required).
3.  **Step 1**: Select your file (Excel/CSV) or Copy data to Clipboard.
4.  **Step 2**: Choose where to save images.
5.  **Step 3**: Confirm the columns (Auto-detected).
6.  Click **Start**.

### For Developers (源码运行)

```bash
# 1. Clone the repo
git clone [https://github.com/youngoris/SheetPic.git](https://github.com/youngoris/SheetPic.git)
cd SheetPic

# 2. Install dependencies
pip install pandas openpyxl xlrd lxml requests pillow pyinstaller

# 3. Run
python sheetpic_v5.py

# SheetPic 📊➡️🖼️

**The smartest way to batch extract and rename images from Excel.**

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Made%20with-Python-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS-lightgrey.svg)]()

## 🧐 Why SheetPic?

If you work in e-commerce, retail, or data management, you've probably faced this nightmare: You have an Excel sheet with hundreds of product photos, and you need to save them all as files named after their Barcodes (SKU).

**Traditional methods fail:**
* ❌ **Right-click Save:** Impossible for 500+ images.
* ❌ **VBA Macros:** Often rely on "Screenshots" (`Chart.Export`), resulting in **blurry, low-res images** or white borders.
* ❌ **Zip Extraction:** Gets you the images, but the filenames are random (`image1.jpg`, `image2.jpg`) and don't match your data.

**SheetPic solves this.** It directly reads the Excel file structure to extract the **original, lossless image file** and matches it with your specified column (e.g., Barcode) instantly.

## ✨ Features

* **100% Lossless Quality**: Extracts the exact binary file stored in the Excel sheet.
* **Smart Renaming**: Automatically names files based on any column you choose (Barcodes, Names, IDs).
* **Intelligent UI**:
    * Auto-detects which columns contain images.
    * Auto-detects header rows to suggest the correct naming column.
* **Safe & Clean**: Prevents desktop clutter by automatically creating a sub-folder for your images.
* **No Code Required**: Comes with a user-friendly Graphical Interface (GUI).

## 📥 Download & Usage (For Users)

1.  Go to the **[Releases](../../releases)** page (Link to your release).
2.  Download `SheetPic.exe`.
3.  **Run the app**:
    * **Step 1**: Select your Excel file (`.xlsx` or `.xlsm`).
    * **Step 2**: Choose where to save the images (Default is Desktop).
    * **Step 3**: Confirm the columns (e.g., Image is in Column D, Barcode is in Column E).
    * **Step 4**: Click **Start Export**.

## 💻 Development (For Developers)

If you want to run the source code or build it yourself.

### Prerequisites
* Python 3.10+
* Pip

### Installation

```bash
git clone [https://github.com/yourusername/SheetPic.git](https://github.com/yourusername/SheetPic.git)
cd SheetPic
pip install -r requirements.txt

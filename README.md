# LibreOffice Calc Image Loader

## Purpose

Load an image into a spreadsheet, where a pixel is represented by a cell.

## Requirements

LibreOffice needs to have Python script support provider:

```
apt install libreoffice-script-provider-python
```

The script itself needs Pillow:

```
pip install Pillow
```

## Usage

Copy `load_image.py` to your LibreOffice Python scripts directory (`~/.config/libreoffice/4/user/Scripts/python`).

Fire up LibreOffice Calc, enter full path of image file in cell A1, execute LoadImage, and wait patiently. Repeatedly setting CellBackColor is slow.

![Demo](demo.gif)

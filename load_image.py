# Requires Pillow to be installed:
# pip install Pillow
from PIL import Image


def LoadImage():
    """Load an image into a spreadsheet, where a pixel is represented by a cell.
    Usage: Put full path of image file in cell A1, execute LoadImage, and wait patiently.
    Repeatedly setting CellBackColor is slow."""

    # Image will be scaled to this many pixels along the longer edge (width or height).
    LONGER_EDGE_PIXELS = 100

    # May need adjusting.
    COLUMN_WIDTH = 200
    ROW_HEIGHT = 180

    def _rgb(r, g, b):
        return (r << 16) | (g << 8) | b

    # Assuming a spreadsheet is currently open.
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    active_sheet = model.CurrentController.ActiveSheet

    path = active_sheet.getCellByPosition(0, 0).String

    try:
        im = Image.open(path)
        scale_factor = LONGER_EDGE_PIXELS / max(im.width, im.height)
        out_w = int(scale_factor * im.width)
        out_h = int(scale_factor * im.height)
        out = Image.new("RGB", (out_w, out_h), (255, 255, 255))
        rsz = im.resize((out_w, out_h))
        out.paste(rsz)
        image_load_successful = True
    except Exception as ex:
        active_sheet.getCellByPosition(0, 0).String = str(ex)
        image_load_successful = False

    if not image_load_successful:
        return None

    active_sheet.getCellByPosition(0, 0).String = ""

    model.enableAutomaticCalculation(False)
    model.lockControllers()
    model.addActionLock()

    columns = active_sheet.getColumns()
    rows = active_sheet.getRows()
    for y in range(out_h):
        rows[y].Height = ROW_HEIGHT
    for x in range(out_w):
        columns[x].Width = COLUMN_WIDTH

    for y in range(out_h):
        for x in range(out_w):
            active_sheet.getCellByPosition(x, y).CellBackColor = _rgb(*out.getpixel((x, y)))

    model.removeActionLock()
    model.unlockControllers()
    model.enableAutomaticCalculation(True)

    return None

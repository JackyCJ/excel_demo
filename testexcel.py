import os
import win32com.client
from PIL import ImageGrab

ExlList = {}
app = None


def Exl_Open(file):
    global app
    app = win32com.client.Dispatch("Excel.Application")
    app.EnableEvents = False
    app.DisplayAlerts = False
    Index = len(ExlList)
    ExlList[Index] = {}
    ExlList[Index]["book"] = app.Workbooks.Open(file)
    return Index


def Exl_Close(id):
    if ExlList[id] != None:
        ExlList[id]["book"].Close()


def Exl_Capture(id, sheet, screen_area, imgpath):
    if ExlList[id] != None:
        ExlList[id]["sheet"] = ExlList[id]["book"].Worksheets(sheet)
        ExlList[id]["sheet"].Range(screen_area).CopyPicture()
        ExlList[id]["sheet"].Paste(ExlList[id]["sheet"].Range("A1"))
        # print(app.Selection.ShapeRange.Name)
        ExlList[id]["sheet"].Shapes(app.Selection.ShapeRange.Name).Copy()
        img = ImageGrab.grabclipboard()
        img.save(imgpath)


if __name__ == '__main__':
    current_path = os.getcwd()
    # xlsx
    id = Exl_Open(current_path + "/data/test.xlsx")
    Exl_Capture(id, "Sheet1", "A1:F8", current_path + "/data/xlsx_capture1.png")
    Exl_Capture(id, "Sheet2", "A1:G8", current_path + "/data/xlsx_capture2.png")
    Exl_Close(id)
    # xls
    id = Exl_Open(current_path + "/data/test.xls")
    Exl_Capture(id, "Sheet1", "A1:F8", current_path + "/data/xls_capture1.png")
    Exl_Capture(id, "Sheet2", "A1:G8", current_path + "/data/xls_capture2.png")
    Exl_Close(id)

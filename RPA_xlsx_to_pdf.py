import os
import pathlib
import win32com.client


def all_files(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


if __name__ == '__main__':
    s = input("Dir: ")
    root_dir = s.strip('\"')

    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False

    for i in all_files(root_dir):
        xlsx = pathlib.Path(i)
        if xlsx.suffix == ".xlsx":
            print(i)
            xlsx_dir = xlsx.parent
            xlsx_dir = str(xlsx_dir)
            basename = xlsx.stem
            basename = str(basename)
            output_file = xlsx_dir + "/" + basename + ".pdf"
            book = app.Workbooks.Open(xlsx)
            xlTypePDF = 0
            book.ActiveSheet.ExportAsFixedFormat(xlTypePDF, output_file)


    app.Quit()

    print("\nDone!")
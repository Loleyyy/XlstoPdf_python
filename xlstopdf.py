import win32com.client
from pywintypes import com_error


# Path to original excel file
WB_PATH = r'C:\Users\loleyy\Documents\Taff_python\test.xlsx'
# PDF path when saving
PATH_TO_PDF = r'C:\Users\loleyy\Documents\Taff_python\test.pdf'

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
try:
    print('Start conversion to PDF')
    wb = excel.Workbooks.Open(WB_PATH)
    # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
    ws_index_list = [1]
    wb.WorkSheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)

except com_error as e:
    print('failed.')
else:
    print('Succeeded.')
finally:
    wb.Close()
    excel.Quit()

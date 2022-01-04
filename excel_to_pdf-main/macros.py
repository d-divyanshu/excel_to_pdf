import os
import pandas
from pickle import TRUE
import win32com.client
from win32com.client import constants as c


# Convert excel to pdf
def excel_to_pdf(excel_file, pdf_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.EnableEvents = True
    # Disable protected mode
    excel.DisplayAlerts = False

    # Disable macros prompts
    excel.AskToUpdateLinks = True
    wb = excel.Workbooks.Open(excel_file)

    #get all sheets
    x=1
    df = pandas.ExcelFile(excel_file)
    ws_index_list = [1]
    bf = df.sheet_names
    for y in bf:
        ws_index_list.append(x)
        x+=1
        print(y)
    ws_index_list.pop()
    

    #scale each sheet
    # print_area = 'A1:BG50'
    for index in ws_index_list:
        ws = wb.Worksheets[index - 1]
        ws.PageSetup.Zoom = False
        # ws.PageSetup.PrintArea = print_area
        #ws.PageSetup.FitToPagesTall = 1
        ws.PageSetup.FitToPagesWide = 1
        

    #merge the sheets

    wb.WorkSheets(ws_index_list).Select()

    #convert to pdf
    wb.ActiveSheet.ExportAsFixedFormat(0, pdf_file)
    try:
        pass
        wb.ActiveSheet.UsedRange
    except Exception as e:
        print(e)
    finally:
        wb.Close()
        excel.Quit()


excel_to_pdf(os.path.abspath('BOQ_681666.xls'),
             os.path.abspath('BOQ_681666.pdf'))

from flask import Flask, request, redirect, url_for, render_template, send_from_directory
from werkzeug.utils import secure_filename
from PyPDF2 import PdfFileWriter
import os
import pandas
from pickle import TRUE
import win32com.client
from win32com.client import constants as c

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
ALLOWED_EXTENSIONS = {'pdf', 'xls'}

app = Flask(__name__, static_url_path="/static")
DIR_PATH = os.path.dirname(os.path.realpath(__file__))
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
# limit upload size upto 8mb
app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            print('No file attached in request')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            # excel_to_pdf(os.path.abspath(file), os.path.abspath('xyz.pdf'))
            # return redirect(url_for('uploaded_file', filename='xyz.pdf'))
    return render_template('index.html')  
 
def excel_to_pdf(excel_file, pdf_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.EnableEvents = True
    # Disable protected mode
    excel.DisplayAlerts = True

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


@app.route('/uploads/<filename>')
def uploaded_file(filename):
   return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
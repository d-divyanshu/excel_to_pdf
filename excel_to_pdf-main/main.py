from flask import Flask, request, redirect, url_for, render_template, send_file
import os
from pickle import TRUE
import win32com.client
import shutil
import uuid
import pythoncom
import PyPDF2

UPLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/uploads/'
DOWNLOAD_FOLDER = os.path.dirname(os.path.abspath(__file__)) + '/downloads/'
ALLOWED_EXTENSIONS = {'pdf', 'xls', 'xlsx', 'xlsb', 'xltx', 'xlsm', 'xltm', 'xlt', 'xml', 'xlam', 'xla', 'xlw', 'xlr', 'csv'}

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
        file.save(os.path.join(app.config['UPLOAD_FOLDER'],file.filename))
        if file.filename == '':
            print('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # filename = secure_filename(file.filename)
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            path_pdf = excel_to_pdf(excel,os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'],file.filename)))
            excel.Quit()
            #os.remove(os.path.abspath(os.path.join(app.config['UPLOAD_FOLDER'],file.filename)))
            # return redirect(url_for('uploaded_file', filename='result.pdf'))
    return render_template('index.html')  


def excel_to_pdf(excel, excel_file):
    excel.Visible = True
    excel.EnableEvents = True
    wb = excel.Workbooks.Open(excel_file, UpdateLinks=0)
    conversion_done = False
    temp_dir = os.path.abspath(os.path.join('temp', uuid.uuid4().hex))
    output_dir = os.path.abspath('output')

    sheets = list(wb.Sheets)
    try:
        for sheet in sheets:
            if sheet.Name == 'Macros':
                continue
            sheet.Activate()
            active_sheet = wb.ActiveSheet

            # Setup the page
            active_sheet.PageSetup.LeftMargin = 0
            active_sheet.PageSetup.RightMargin = 0
            active_sheet.PageSetup.TopMargin = 0
            active_sheet.PageSetup.BottomMargin = 0
            active_sheet.PageSetup.HeaderMargin = 0
            active_sheet.PageSetup.FooterMargin = 0
            active_sheet.PageSetup.CenterHorizontally = True

            # Fit to page wide
            active_sheet.PageSetup.Zoom = False
            active_sheet.PageSetup.FitToPagesTall = False
            active_sheet.PageSetup.FitToPagesWide = 1

            os.makedirs(temp_dir, exist_ok=True)
            output_file = os.path.join(
                temp_dir, f'{sheets.index(sheet)} - {sheet.Name}.pdf')
            try:
                print(f'Converting sheet: {active_sheet.Name}')
                wb.ActiveSheet.ExportAsFixedFormat(0, output_file)
                conversion_done = True
            except Exception:
                continue
    except Exception as e:
        print(e)
    finally:
        excel.EnableEvents = False
        excel.DisplayAlerts = False
        wb.Close()

    # Return if conversion was not done
    if not conversion_done:
        return

    # Merge all the pdfs
    pdf_merger = PyPDF2.PdfFileMerger()
    for filename in os.listdir(temp_dir):
        if filename.endswith('.pdf'):
            pdf_merger.append(os.path.join(temp_dir, filename))

    # Save the merged pdf
    os.makedirs(output_dir, exist_ok=True)
    final_pdf_file = os.path.join(
        output_dir, os.path.basename(excel_file).split('.')[0] + '.pdf')
    with open(final_pdf_file, 'wb') as fobj:
        pdf_merger.write(fobj)
    pdf_merger.close()

    # Remove the temp directory
    shutil.rmtree(temp_dir)
    return final_pdf_file


@app.route('/<filename>')
def uploaded_file(filename):
   return send_file( os.path.abspath('result.pdf'), filename, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
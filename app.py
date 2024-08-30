# 사이트에 올리는 거 필요함,,,!
from flask import Flask, render_template, request, send_file
import openpyxl
from vobject import vCard
import io

app = Flask(__name__)

def create_vcard(full_name, phone_number):
    card = vCard()
    card.add('fn').value = full_name
    card.add('tel').value = phone_number
    return card

def excel_to_vcards(excel_file):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    
    output = io.StringIO()
    for row in sheet.iter_rows(min_row=2, values_only=True):
        full_name, phone_number = row[0], row[1]
        vcard = create_vcard(full_name, phone_number)
        output.write(vcard.serialize())
    
    return output.getvalue()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            vcf_content = excel_to_vcards(file)
            output = io.BytesIO()
            output.write(vcf_content.encode('utf-8'))
            output.seek(0)
            return send_file(output, mimetype='text/vcard',
                             as_attachment=True, download_name='contacts.vcf')
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
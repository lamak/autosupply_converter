import os
import uuid

from dotenv import load_dotenv
from flask import Flask, flash, redirect, url_for, send_from_directory, render_template
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from flask_wtf.file import FileRequired, FileAllowed
from werkzeug.utils import secure_filename
from wtforms import FileField

from converter import process_supply

load_dotenv()
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev')
bootstrap = Bootstrap(app)

PATH_UPLOAD = os.environ.get('PATH_UPLOAD', 'uploads')
PATH_RESULT = os.environ.get('PATH_RESULT', 'results')
PATH_EXPORT = os.environ.get('PATH_EXPORT', '.')
TEMPLATE_FILE = os.environ.get('TEMPLATE_FILE', 'sku-body-template.xlsx')


class FileUploadForm(FlaskForm):
    upload = FileField(validators=[FileRequired(), FileAllowed(upload_set=['xlsx', 'docx'])])


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    form = FileUploadForm()
    params = {
        'template_name_or_list': 'index.html',
        'title': 'Конвертер остатков автозаказа',
        'form': form,
    }
    if form.validate_on_submit():
        f = form.upload.data
        filename = f'{uuid.uuid4()}_{secure_filename(f.filename)}'
        filepath = os.path.join(PATH_UPLOAD, filename)
        f.save(filepath)
        errors = list()

        result, errors = process_supply(filepath, PATH_EXPORT, PATH_RESULT, TEMPLATE_FILE)

        if errors:
            flash('\n'.join(errors))

        if result:
            return redirect(url_for('download_result', filename=result))

    return render_template(**params)


@app.route('/results/<filename>')
def download_result(filename):
    return send_from_directory(PATH_RESULT, filename)

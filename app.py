import os
import uuid

from flask import Flask, flash, request, redirect, url_for, send_from_directory, render_template
from flask_bootstrap import Bootstrap
from flask_wtf import FlaskForm
from werkzeug.utils import secure_filename
from wtforms import FileField

from converter import process_supply

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY')
bootstrap = Bootstrap(app)


class FileUploadForm(FlaskForm):
    file = FileField()


def validate_filename(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'xlsx'


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    params = {
        'template_name_or_list': 'index.html',
        'title': 'Конвертер остатков автозаказа',
        'form': FileUploadForm(),
    }

    if request.method == 'POST':
        if not request.files.get('file'):
            flash('Файл не выбран')
            return redirect(request.url)

        file = request.files['file']

        if not validate_filename(file.filename):
            flash('Выберите XLSX документ')
            return redirect(request.url)

        if file:
            filename = f'{uuid.uuid4()}_{secure_filename(file.filename)}'
            filepath = os.path.join(os.environ.get('UPLOAD_FOLDER'), filename)
            file.save(filepath)
            errors = list()

            result, errors = process_supply(filepath)

            if errors:
                flash('\n'.join(errors))

            if result:
                return redirect(url_for('download_result', filename=result))

    return render_template(**params)


@app.route('/results/<filename>')
def download_result(filename):
    return send_from_directory('results', filename)

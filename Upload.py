#Importar bibliotecas necessárias
from flask import Flask, flash, request, redirect, url_for, session, render_template, send_from_directory
from werkzeug.utils import secure_filename
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from werkzeug.exceptions import RequestEntityTooLarge
import os
from beckendUploadFiles import *


app = Flask(__name__)

#Regras e configurações
UPLOAD_FOLDER = 'Uploads'
ALLOWED_EXTENSIONS =  {'pdf', 'csv', 'doc', 'docx', 'txt','xlsx', 'xls'}
SIZE = 16 * 1000 * 1000

app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['ALLOWED_EXTENSIONS'] = ALLOWED_EXTENSIONS
app.config['MAX_CONTENT_LENGTH'] = SIZE

class UploadFileForm(FlaskForm):
    file = FileField("File", validators=[InputRequired()])
    submit = SubmitField("Upload File")

#Validar as extensões dos arquivos
def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


#Avalia o arquivo importado, extensão, salva o arquivo na pasta Uploads e chama a função save_index que fará 
#o uploado no elasticsearch

#determinar a rota de acesso
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    form = UploadFileForm()
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Falta parte do arquivo')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('Nenhum arquivo selecionado')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)            
            file.save(os.path.join(os.path.abspath(os.path.dirname(__file__)), app.config['UPLOAD_FOLDER'], secure_filename(file.filename)))
            save_index(r'C:\Users\adria\Documents\Trabalho - Central\codigos\Novapasta\UploadFiles\Uploads')
            flash('Arquivo importado com sucesso!')
            return redirect(request.url)
        
        else:
            flash("Erro ao importar, tente mais tarde!")                 
    return render_template("upload.html", form=form)



if __name__=="__main__":
    app.run(debug=True)
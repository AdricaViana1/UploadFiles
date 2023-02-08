#Importar bibliotecas necessárias
from flask import Flask, flash, request, redirect, url_for, session, render_template, send_from_directory
from werkzeug.utils import secure_filename
from flask_wtf import FlaskForm
from wtforms import FileField, SubmitField
from werkzeug.exceptions import RequestEntityTooLarge
from wtforms.validators import InputRequired
from elasticsearch import Elasticsearch, helpers, ElasticsearchException
from datetime import datetime
from glob import glob
from dataclasses import dataclass, field
from typing import List
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from xml.dom import NotFoundErr
from json import JSONEncoder
from bs4 import BeautifulSoup
from docx2python import docx2python
import aspose.words as aw
import os
import requests
import urllib3
import json
import pandas as pd
import PyPDF2
import re
import xlrd
import pandas as pd
import simplejson as json

#Parametros de conexão com Elasticsearch
es = Elasticsearch(
        hosts="https://localhost:9200",
        ca_certs="C:/Users/adria/Documents/Trabalho - Central/codigos/http_ca.crt",
        http_auth=("elastic", "WTDxSVLE0Wx-EK4Ngoh="),
        verify_certs=True
        )

#Função que realiza a limpeza da pasta de Uploads. Executada após a importação de cada arquivo.       
def clear():
    dir = r'C:\Users\adria\Documents\Trabalho - Central\codigos\Novapasta\UploadFiles\Uploads'
    for limpar in os.listdir(dir):
        os.remove(os.path.join(dir, limpar))
        print("Arquivos removidos")

#cria uma lista vazia, com o for lê o nome da pasta raiz, pasta atual e do arquivo e salva na lista vazia.
def buscar_arquivos(caminho):
    lista_arquivos = []
    
    for raiz, diretorios, arquivos in os.walk(caminho):
       for arquivo in arquivos:
             lista_arquivos.append(os.path.join(raiz, arquivo))
    return lista_arquivos


#Função para busca o nome dos arquivos que estão na lista da função buscar_arquivos, avalia o que existe no nome do último arquivo após o .  
# Uma condição para tratamento de cada formato e upload do arquivo no elasticsearch
def save_index(file):
    arquivos = buscar_arquivos(file)
    arquivo = str(arquivos[0]).split(".")[-1]    
    
    #.CSV
    if arquivo == 'csv':
        with open(arquivos[0], 'r', encoding='latin-1') as fileopen:        
            Result = fileopen.read()
            
            #converti em dataframe para converter para Json
            action = pd.DataFrame([Result]) 
            
            #converter para dicionário
            arquivo = action.to_dict('list')
            
            #converter para Json
            import json
            df = json.dumps(arquivo, ensure_ascii=False)
            
            #salvar no elasticsearch
            res = es.index(index='import_files', document=df) #Salvar no Elasticsearch
            print("\nArquivo .csv importado\n", res)
        clear()         

    #.PDF                                             
    if arquivo == 'pdf':
        with open(arquivos[0], 'rb') as fileopen:
            dados_pdf = PyPDF2.PdfFileReader(fileopen)
            
            output = ''

            count = dados_pdf.numPages
            #pegar todas as páginas do arquivo e extrair todo o texto.
            for i in range(count):
                page = dados_pdf.getPage(i)
                output += page.extractText()

            #subistituir quebra de linha
            texto1 = re.sub('\n', ' ', output)
            texto1 = re.sub('\n', '', texto1)
            
            #coloca um titulo antes do texto
            titulo = texto1.split()
            titulo.insert(0, "Dados.: ")
            final = ' '.join(titulo)
            #determinar a separação do titulo e do texto 
            u = final.split(".: ")
                
            #Converter em dicionário
            def extrai_pdf(a):
                it = iter(a)
                res = dict(zip(it, it))
                return res
            
            import json
            dom = json.dumps(extrai_pdf(u),ensure_ascii=False, separators=(",", ":"))
                    
            #Inserir o documento no elasticsearch
            res = es.index(index="import_files", document=dom)
                        
            print("\n1. Arquivo PDF importado\n", res)
            fileopen.close()
        
        clear()
    #.TXT    
    if arquivo == 'txt':
        with open(arquivos[0], 'r', encoding='latin-1') as txt_file:
            txt_file = txt_file.read()
            
            #ler com o pandas para que tenha o formato de dataframe
            txt_data = pd.DataFrame([txt_file]) 

            #Converter para dicionario de lista
            txt_data_ = txt_data.to_dict('list')
            import json
            Dict_txt = json.dumps(txt_data_, ensure_ascii=False, separators=(",", ":"))
            
            res = es.index(index="import_files", document=Dict_txt)
            print("\nArquivo .txt importado\n", res)
        clear()
    #.XLSX
    if arquivo == 'xlsx':
        data = xlrd.open_workbook(arquivos[0])
        
        #Ler a primeira aba
        for i in range(data.sheet_by_index(0).nrows):
            print(data.sheet_by_index(0).row_values(i))           
          
        dict_df = {}
        #Função para Ler os dados importados e inseri-los como string no dicionário dict_df
        def string_to_dict(data):   
            for sheet in data.sheet_names():
                sheet_data = []
                for i in range(data.sheet_by_index(0).nrows):
                    sheet_data.append(data.sheet_by_index(0).row_values(i))
                    dict_df[sheet] = str(sheet_data)
                
                print(dict_df)
        
        #Executar a função    
        string_to_dict(data)    
            
        import json
        
        #Converter para o formato de string json
        _xlsx = json.dumps(dict_df,ensure_ascii=False, separators=(",", ":"))         
        
        #Importar para o elastic        
        res = es.index(index="import_files",document=_xlsx)
        
        #Plota os dadps do index     
        print("\nArquivo .XLSX importado\n", res)
        
        #Limpa os arquivos da pasta de Uploads    
        clear()
        
    #.XLS
    if arquivo == 'xls':
        book = xlrd.open_workbook(arquivos[0])
        plan = book.sheets()[0] 
       
        for i in range(book.sheet_by_index(0).nrows):
            print(book.sheet_by_index(0).row_values(i))
            
        #coletar os dados do arquivo, como string e colocar dentro do dicionário    
        dic_df = {}
        
        def xls_to_dict(book):
            for sheet in book.sheet_names():
                sheet_data = []
                for i in range(book.sheet_by_index(0).nrows):
                    sheet_data.append(book.sheet_by_index(0).row_values(i))
                    
                    dic_df[sheet] = str(sheet_data)
            
        xls_to_dict(book)
                                
        #Converter em json
        import json
                
        _xls = json.dumps(dic_df,ensure_ascii=False, separators=(",", ":") )
         
        res = es.index(index="import_files", document=_xls)
        print("\nArquivo .XLS importado\n", res)
        clear()
        
    #.DOCX
    if arquivo == 'docx':
        with open(arquivos[0], 'rb') as docx_file:
            data_doc = docx2python(docx_file) # armazena o conteúdo recuperado de dentro do arquivo    
            doc = data_doc.body
            doc1 = doc[0]
            doc2 = doc1[0]
            doc3 = doc2[0]

            #excluir '' 
            while('' in doc3):
                doc3.remove('')
                
                #excluir '' 
            while('\t' in doc3):
                doc3.remove('\t')

            arquivoDoc1 = list(doc3)

            listaw1 = arquivoDoc1[:-1:2]
            listaw2 = arquivoDoc1[1::2]

            resw = dict(zip(listaw1, listaw2))
            import json
            new_string = json.dumps(resw, ensure_ascii=False, separators=(",", ":"))
            
            print(new_string)
                            
            res = es.index(index="import_files", document=new_string)
                
            print("\nArquivo .DOCX importado\n", res)      
        clear()
        
    #.DOC              
    if arquivo == 'doc':
        document = aw.Document(arquivos[0])
        builder = aw.DocumentBuilder(document)

        #Salvei como txt
        document.save("doc_convertido.txt")

        #Ler o arquivo Txt retirando as quebras de linhas
        with open("doc_convertido.txt", "r", encoding='utf-8') as outfile:
            convertido = [linha.strip() for linha in outfile if linha.strip() != ""]

            #Excluir  as informações da biblioteca nas primeiras linhas e últimas linhas -1, -2 
            del convertido[0]
            del convertido[-1]
            del convertido[0]
            del convertido[-1]

            #Deixar no formato de dicionario chave - valor
            def extrai_doc(a):
                it = iter(a)
                res = dict(zip(it, it))
                return res

            extrai_doc(convertido)

            #Converter para Json file
            import json
            conv = json.dumps(extrai_doc(convertido), ensure_ascii=False)
            
            res = es.index(index="import_files", document=conv)
            
            print("Arquivo .doc importado com sucesso", res)
        clear()


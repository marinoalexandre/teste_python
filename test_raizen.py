#__name__    = 'test_raizen'
#__author__  = 'Alexandre Marino'
#__date__    = '2021-12-17'
#__version__ = '1.0'

import os, sys
import zipfile
import numpy as np
import pandas as pd
import xml.etree.ElementTree as et
import win32com.client as win32
from datetime import datetime

#global 
#o arquivo de origem dos dados deve estar na pasta: exame_raizen/data/vendas-combustiveis-m3.xls
fname     = 'vendas-combustiveis-m3'
fpathini  = 'exame_raizen/'
fpathdata = fpathini + 'data/'
fpathout  = fpathini + 'output/'
fpathzip  = fpathini + 'unzip/'
fpath     = fpathzip + 'xl/pivotCache/'

def process_xml(fdef,frec,foutput):
    
    def check_dir(dir): 
        if not os.path.exists(dir):
            os.makedirs(dir)    

    def conv_to_xlsx(): #convert xls format to xlsx function
        try:
            file_name = fpathdata + fname + '.xls'
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(file_name)     
            wb.SaveAs(file_name+"x", FileFormat = 51)
            wb.Close()                               
            excel.Application.Quit()
            print('Arquivo convertido de xls para xlsx.')
        except:
            print('Falha na conversão de formato')
            return None
    
    def rename_xlsx_to_zip(): #altera a extenção do arquivo de xlsx para zip
        try:
            dir = os.listdir(fpathdata)
            for file in dir:
                if file == fname:
                    os.rename(fname + '.xlsx', fname + '.zip')
                    os.remove(fname + '.xlsx')
                    print('A extensão do arquivo', file, 'foi alterada de xlsx para zip')
        except:
            print('Falha no rename')
            return None

    def unzip_file(): #cria diretório unzip e descompacta zip
        check_dir(fpathzip)
        try:
            with zipfile.ZipFile(fpathdata + fname + '.zip', 'r') as zip_ref:
                 zip_ref.extractall(fpathzip)
            os.remove(fpathdata + fname + '.zip')
        except:
            print('Falha no unzip')
            return None

    print('\nConvertendo xls em xlsx...')
    conv_to_xlsx()
    
    print('\nRenomeando para zip...')
    rename_xlsx_to_zip()
    
    print('\nUnziping...')
    unzip_file()

    definitions = fpath + fdef
    print('\nDefinition path:',definitions)  

    # dicionario e colunas da pivot 
    defdict = {}
    columnas = []
    try:
        re = et.parse(definitions).getroot()
    except:
        print('Arquivo',definitions,'não encontrado!')
        return None
    
    for fields in re.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheFields'):
        for cidx, field in enumerate(list(fields)):
            columna = field.attrib.get('name')
            defdict[cidx] = []
            columnas.append(columna)
            for value in list(list(field)[0]):
                tagname = value.tag
                defdict[cidx].append(value.attrib.get('v', 0))

    # recupera os dados da pivot records   
    dfdata = [] 
    bdata = fpath + frec
    print('Records path:',bdata)

    try:
        for event, elem in et.iterparse(bdata, events=('start', 'end')):
            if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r' and event == 'start':
                tmpdata = []
                for cidx, valueobj in enumerate(list(elem)): 
                    tagname = valueobj.tag
                    vattrib = valueobj.attrib.get('v')
                    rdata = vattrib

                    if tagname == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x': 
                        try:
                            rdata = defdict[cidx][int(vattrib)]
                        except:
                            logging.exception("message")

                    tmpdata.append(rdata)
                if tmpdata:
                    dfdata.append(tmpdata)
                elem.clear()

    except:
        print('Arquivo',bdata,'não encontrado!')
        return None

    #dataframe
    df = pd.DataFrame(dfdata)
    df.columns = ['product', 'year', 'region', 'uf', 'unit', '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', 'total'] #coloca nome nas colunas
    df = df.drop('total', 1)   #exclui coluna "total"

    for column in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']: 
        df[column] = df[column].astype(float)

    df = df.melt(id_vars=['product','year','region','uf','unit'], var_name=['month'], value_name='volume') # pivot invertida
    df['year_month'] = df['year'] + '-' + df['month']  # adiciona a coluna no df concatenando ano e mês
    df = df.drop(['year', 'region', 'month'], 1)       # exclui colunas desnecessárias
    df['created_at'] = datetime.now()                  # insere nova coluna com conteudo de data_atual
    df = df[['year_month', 'uf', 'product', 'unit', 'volume', 'created_at']] # organiza colunas conforme schema definido

    #output result
    check_dir(fpathout)
    
    #------ salva dados em csv (para visualizar os dados no excel)
    #df.to_csv(fpathout + foutput + '.csv', sep=';', header='true', index_label='id', encoding='utf8') 

    #------ salva dados em parquet (escolhido para o test)
    df.to_parquet(fpathout + foutput + '.parquet', engine='auto')
    
    #------ salva dados em parquet particionado (performatico)
    #df.to_parquet(fpathout + foutput, engine='auto', partition_cols=['year_month'])
    #end def
    
def print_validation():
    # read parquet
    print('\nValidação dos dados: lendo parquet... ')
    pd.options.display.float_format = '$ {:,.0f}'.format
    
    df = pd.read_parquet(fpathout + 'fuel_uf.parquet')
    table = pd.pivot_table(df, values='volume', index='year_month', aggfunc=np.sum)
    #table.query('year_month == ["2013-01"]')

    print('\n\nSales of oil derivative fuels by UF and product (2000-Jan):')
    print(table.head(1))
    
    df = pd.read_parquet(fpathout + 'diesel_uf.parquet')
    table = pd.pivot_table(df, values='volume', index='year_month', aggfunc=np.sum)
    
    print('\n\nSales of diesel by UF and type (2013-Jan):')
    print(table.head(1))
    

process_xml('pivotCacheDefinition1.xml','pivotCacheRecords1.xml','fuel_uf')
process_xml('pivotCacheDefinition2.xml','pivotCacheRecords2.xml','diesel_uf')

print_validation()


#################################################################
#Validação dos dados: lendo parquet... 
#
#
#Sales of oil derivative fuels by UF and product (2000-Jan):
#                volume
#year_month            
#2000-01    $ 6,991,363
#
#
#Sales of diesel by UF and type (2013-Jan):
#                volume
#year_month            
#2013-01    $ 4,456,693

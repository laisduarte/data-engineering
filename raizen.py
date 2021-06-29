import pandas as pd
from datetime import datetime
import win32com.client as win32
from pathlib import Path
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sqlalchemy


def extract (tables):
    newDf = pd.DataFrame()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = True
    wb = excel.Workbooks.Open(Path.cwd() / "vendas-combustiveis-m3.xls")
    
    def setFilter(pvtTable, filterName, value):
        pvtTable.PivotFields(filterName).ClearAllFilters()
        try:
            pvtTable.PivotFields(filterName).CurrentPage = str(value.Name)
        except:
            for item in pvtTable.PivotFields(filterName).PivotItems():
                if item.Name == value.Name:
                    item.Visible = True
                else:
                    item.Visible = False
    
    for table in tables:
        pvtTable = wb.Sheets("Plan1").Range(table).PivotTable
        
        for uf in pvtTable.PivotFields("UN. DA FEDERAÇÃO").PivotItems():  
            setFilter(pvtTable, "UN. DA FEDERAÇÃO", uf)
            
            for product in pvtTable.PivotFields("PRODUTO").PivotItems():
                setFilter(pvtTable, "PRODUTO", product)
                
                arrAux = []
                arrComplete = []
                curLine = pvtTable.TableRange1[0].Row
                
                for item in pvtTable.TableRange1:
                    if item.Row == curLine:
                        arrAux.append (item.Value)
                    else:
                        arrComplete.append (arrAux)
                        arrAux = [];
                        arrAux.append (item.Value)
                    
                    curLine = item.Row
                
                for year in arrComplete[1][1:]:
                    for i in range(2,len(arrComplete),1):
                        newDf = newDf.append (pd.DataFrame({'year_month':[datetime(int(year), int(i-1), 1)],'uf': [str(uf)], 'product': [str(product)], 'unit' : ['m3'], 'volume' : [arrComplete[i][arrComplete[1].index(year)]], 'created_at' : [datetime.today()]}))

        
    newDf=newDf.astype({'uf':str,'product':str,'unit':str,'volume':float })        
    return newDf


def generalCheck (newDf):
    print("Abaixo serão exibidos os valores únicos de cada uma das colunas do dataframe. O resultado esperado são 27 unidades federativas e 13 products (8 opções da primeira tabela e 5 da segunda). Analisando o dataframe criado temos: ")
    print(newDf.nunique())
    print("-------------------------------------------------------------")     
    print("Considerando as variáveis que temos, são esperadas 67392 linhas na tabela (fora o título): ")
    print("Tabela 1 + Tabela 2")    
    print("[27(uf) * 8 (product) * 12 (months) * 21 (years)] + [27(uf) * 5 (product) * 12 (months) * 8 (years)]") 
    print("54432 + 12960 = 67392") 
    print (len(newDf.index))
    print("-------------------------------------------------------------")
    print("Espera-se também uma contagem homogênea de valores em cada uma das unidades federativas, ou seja (67392 / 27 = 2496): ")
    print(newDf['uf'].value_counts())
    
def specificCheck(newDf, uf, product, year, month, value):
    volume  = newDf.loc[(newDf['uf'] == uf) & (newDf['product'] == product) & (newDf['year_month'] >= (str(year)+"-"+str(month)+"-01")) & (newDf['year_month'] <= (str(year)+"-"+str(month)+"-01")), 'volume'].sum()
    
    return ("Valor bate com o fornecido" if value == volume else "Valor não bate com o fornecido")

def toGoogleSheets (newDf):
     #link da planilha: https://docs.google.com/spreadsheets/d/1GazgmI2WB6Z-7szTq3WIvYofrM0FadPXhKedk5uxkrE/edit#gid=0
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('keys.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open("Raizen - Data Engineering Test").sheet1
    sheet.clear()
    
    
    #redução do dataframe para teste, por motivos de limitação de cotas de requisição da API
    testDf = newDf.head(20)
    arr = testDf.to_numpy()

    sheet.insert_row(['year_month','uf','product', 'unit','volume', 'created_at'], 1)

    for i in range(len(arr)):
        arrAux = []
        for j in range (len(arr[i])):
            arrAux.append (str(arr[i][j]))
            row = arrAux
            print
            index = i+2
        sheet.insert_row(row, index)

   

def toDatabase(newDf, username, password, ip, databaseName):
    database_username = username
    database_password = password
    database_ip       = ip
    database_name     = databaseName
    
    #excluir
    testDf = newDf.head(20)
    
    database_connection = sqlalchemy.create_engine('mysql+mysqlconnector://{0}:{1}@{2}/{3}'.format(database_username, database_password, database_ip, database_name))
    testDf.to_sql(con=database_connection, name='table_name_for_df', if_exists='replace')

def toCSV (newDf):
    newDf.to_csv('BD.csv', index=False, encoding="iso-8859-1")
    


#####################################################################################################
    
#declara as tabelas dinâmicas que queremos extrair
tables = ["B52","B132"]
#extração dos dados das tabelas
newDf = extract (tables)
#visão geral sobre os dados extraídos
generalCheck(newDf)
#checa o valor fornecido para a função com o existente no dataframe
specificCheck(newDf, "RIO DE JANEIRO", "GLP (m3)", 2002, 3, 83019.42)
#salva dados em formato csv
toCSV (newDf)
#salva da em planilha online através da API do Google Sheets
toGoogleSheets (newDf)
#envia dados para BD MySQL 
toDatabase(newDf, 'root', '', '127.0.0.1', 'anpFuelSale')
        
#####################################################################################################



import subprocess as sp
import sqlite3
from sqlite3 import Error
import requests, json, configparser, datetime
import xml.etree.ElementTree as ET
from urllib3.exceptions import InsecureRequestWarning
# Suppress only the single warning from urllib3 needed.
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

# Archivo de variables de configuracion
configParser = configparser.RawConfigParser()
configFilePath = "./env.cfg"
configParser.read(configFilePath)
# Archivo de variables de configuracion


def rwTxtFile(fileName):
    estadisticas = {}
    estadisticas1 = {}
    with open(fileName,mode='r',encoding='utf-8') as fr:
        lines = fr.read().splitlines()
        for word in lines:
            estadisticas[word.split()[0]] = word.split()[1]
    with open('estadisticas1.txt',mode='r',encoding='utf-8') as fr1:
        lines = fr1.read().splitlines()
        for word in lines:
            estadisticas1[word.split()[0]] = word.split()[1]
    with open('stat.html',mode='w',encoding='utf-8') as f:
        f.write(f'''
<style>
  * {{
    box-sizing: border-box;
  }}
  .row {{
    display: flex;
  }}
  .column {{
    flex: 50%;
    padding: 1px;
  }}
  p {{
    width: 50%;
    border-spacing: 0;
    border-collapse: collapse;
    border: 1px solid#ddd;
    text-align: center;
  }}
  table {{
    border-collapse: collapse;
    border-spacing: 0;
    width: 100%;
    border: 1px solid #ddd;
  }}
  th, td {{
    text-align: left;
    padding: 5px;
  }}
  tr:nth-child(odd) {{
    background-color: #f2f2f2;
}}
</style>
<div class="row">
  <p>Estadisticas (Perdida Neta)
  </p> 
  <p>Estadisticas (Frecuencia)
  </p> 
</div>
<div class="row">
  <div class="column">
    <table>
      <tr>
        <td>Cantidad valores
        </td>
        <td>{estadisticas['nbr.val']}
        </td>
      </tr>
      <tr>
        <td>Cantidad valores nulos
        </td>
        <td>{estadisticas['nbr.null']}
        </td>
      </tr>
      <tr>
        <td>Cantidad de valores perdidos
        </td>
        <td>{estadisticas['nbr.na']}
        </td>
      </tr>
      <tr>
        <td>Valor Minimo
        </td>
        <td>{estadisticas['min']}
        </td>
      </tr>
      <tr>
        <td>Valor Maximo
        </td>
        <td>{estadisticas['max']}
        </td>
      </tr>
      <tr>
        <td>Rango
        </td>
        <td>{estadisticas['range']}
        </td>
      </tr>
      <tr>
        <td>Suma
        </td>
        <td>{estadisticas['sum']}
        </td>
      </tr>
      <tr>
        <td>Mediana
        </td>
        <td>{estadisticas['median']}
        </td>
      </tr>
      <tr>
        <td>Media
        </td>
        <td>{estadisticas['mean']}
        </td>
      </tr>
      <tr>
        <td>Error estándar en la media
        </td>
        <td>{estadisticas['SE.mean']}
        </td>
      </tr>
      <tr>
        <td>Intervalo de confianza de la media 0.95
        </td>
        <td>{estadisticas['CI.mean.0.95']}
        </td>
      </tr>
      <tr>
        <td>Diferencia
        </td>
        <td>{estadisticas['var']}
        </td>
      </tr>
      <tr>
        <td>Desviación estándar
        </td>
        <td>{estadisticas['std.dev']}
        </td>
      </tr>
      <tr>
        <td>Coeficiente de variación
        </td>
        <td>{estadisticas['coef.var']}
        </td>
      </tr>
    </table>
  </div>
  <div class="column">
    <table>
      <tr>
        <td>Cantidad valores
        </td>
        <td>{estadisticas1['nbr.val']}
        </td>
      </tr>
      <tr>
        <td>Cantidad valores nulos
        </td>
        <td>{estadisticas1['nbr.null']}
        </td>
      </tr>
      <tr>
        <td>Cantidad de valores perdidos
        </td>
        <td>{estadisticas1['nbr.na']}
        </td>
      </tr>
      <tr>
        <td>Valor Minimo
        </td>
        <td>{estadisticas1['min']}
        </td>
      </tr>
      <tr>
        <td>Valor Maximo
        </td>
        <td>{estadisticas1['max']}
        </td>
      </tr>
      <tr>
        <td>Rango
        </td>
        <td>{estadisticas1['range']}
        </td>
      </tr>
      <tr>
        <td>Suma
        </td>
        <td>{estadisticas1['sum']}
        </td>
      </tr>
      <tr>
        <td>Mediana
        </td>
        <td>{estadisticas1['median']}
        </td>
      </tr>
      <tr>
        <td>Media
        </td>
        <td>{estadisticas1['mean']}
        </td>
      </tr>
      <tr>
        <td>Error estándar en la media
        </td>
        <td>{estadisticas1['SE.mean']}
        </td>
      </tr>
      <tr>
        <td>Intervalo de confianza de la media 0.95
        </td>
        <td>{estadisticas1['CI.mean.0.95']}
        </td>
      </tr>
      <tr>
        <td>Diferencia
        </td>
        <td>{estadisticas1['var']}
        </td>
      </tr>
      <tr>
        <td>Desviación estándar
        </td>
        <td>{estadisticas1['std.dev']}
        </td>
      </tr>
      <tr>
        <td>Coeficiente de variación
        </td>
        <td>{estadisticas1['coef.var']}
        </td>
      </tr>
    </table>
  </div>
</div>
        ''')


def getHeaders(sessionToken = ''):
    headers = {
        'Accept':'application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Content-Type': 'application/json'
    }

    if sessionToken != '':
        headers['Authorization'] = 'Archer session-id='+sessionToken

    return headers
    
def apiCall(url, headers, content):
    data = content
    headers = headers
    response = requests.post(url, verify=False, headers=headers, json=data)
    return response

class ReporteArcher:
  def __init__(self, id, mes, valor, fecha):
    self.id = id
    self.mes = mes
    self.valor = valor.replace(",",".")
    self.fecha = datetime.datetime.strptime(fecha, "%d/%m/%Y").strftime("%Y-%m-%d")
    

baseurl = configParser.get('env', 'baseUrl')

data = {
    "InstanceName": configParser.get('env', 'instanceName'),
    "Username": configParser.get('env', 'username'),
    "UserDomain": configParser.get('env', 'userDomain'),
    "Password": configParser.get('env', 'password')
}

print(baseurl,data)

sessionToken = apiCall(baseurl+'/api/core/security/login', getHeaders(), data).json()['RequestedObject']['SessionToken']

headers = getHeaders(sessionToken)

userSessionHeader = {'Content-Type': 'text/xml;charset=utf-8','SOAPAction': 'http://archer-tech.com/webservices/SearchRecordsByReport'}

registros = []
i=1
while i>0:
    body = \
            """<?xml version="1.0" encoding="utf-8"?>
               <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
               <soap:Body>
               <SearchRecordsByReport xmlns="http://archer-tech.com/webservices/">
               <sessionToken>"""+sessionToken+"""</sessionToken>
                <reportIdOrGuid>"""+configParser.get('env', 'reportIdOrGuid')+"""</reportIdOrGuid>
                <pageNumber>"""+str(i)+"""</pageNumber>
                </SearchRecordsByReport>
               </soap:Body>
               </soap:Envelope>
            """
    response = requests.post(baseurl + '/ws/search.asmx', verify=False, data=body, headers=userSessionHeader)
    #print response.content
    tree = ET.fromstring(response.content)
    reg = tree.find('.//{http://archer-tech.com/webservices/}SearchRecordsByReportResult').text
    regf = reg.replace("""<?xml version="1.0" encoding="utf-16"?>""", """<?xml version="1.0" encoding="utf-8"?>""")
    regf2 = regf.encode('utf-8', errors='ignore')
    lstUsuarios = ET.fromstring(regf2)
    records = lstUsuarios.findall('Record')
    if not records:
        i=0
    else:
        for record in records:
            registros.append(ReporteArcher(record[0].text, record[1].text, record[2].text, record[3].text))
        i+=1
       

def create_connection(db_file):
    """ create a database connection to a SQLite database """
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)
    return conn
    
def create_table(conn, create_table_sql):
    """ create a table from the create_table_sql statement
    :param conn: Connection object
    :param create_table_sql: a CREATE TABLE statement
    :return:
    """
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)
        
def insert_report(conn, reporte):
    """
    Create a new project into the projects table
    :param conn:
    :param project:
    :return: project id
    """
    sql = ''' REPLACE INTO reportes(id,Mes,PerdidaNeta,FechadeOcurrencia)
              VALUES(?,?,?,?) '''
    cur = conn.cursor()
    cur.execute(sql, reporte)
    conn.commit()
    return cur.lastrowid
        

database = configParser.get('env', 'database')
conn = create_connection(database)
sql_create_reportes_table = """ CREATE TABLE IF NOT EXISTS reportes (
                                        id integer PRIMARY KEY,
                                        Mes text NOT NULL,
                                        PerdidaNeta double,
                                        FechadeOcurrencia DATE NOT NULL
                                    ); """

if conn is not None:
        create_table(conn, sql_create_reportes_table)
        for registro in registros:
            registro_formateado = (registro.id, registro.mes, registro.valor, registro.fecha);
            print(registro_formateado)
            registro_id = insert_report(conn, registro_formateado)
else:
        print("Error! No se pudo conectar a la base de datos")

dir = configParser.get('env', 'dirDeScriptR')
script = dir + configParser.get('env', 'nombreScript')
sp.check_call(['rscript', "--vanilla", script], shell=False)
rwTxtFile('estadisticas.txt')

import pyodbc
import sys
import xmlrpc.client
import os
from datetime import datetime
from openpyxl import Workbook
import xlsxwriter
from datetime import datetime

def mi_proveedor(url, db, uid, password, l_prov):
    import xmlrpc.client
    # Calliing methods
    models_prov = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_prov.execute_kw(db, uid, password,
                     'res.partner', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['name', '=', l_prov.upper()], ['active','=',1]]]  #lista de python
    registros = models_prov.execute_kw(db, uid, password, 'res.partner', 'search_count', filtro)
    ids =       models_prov.execute_kw(db, uid, password, 'res.partner', 'search',       filtro, {'limit': 1})
    if registros == 0:
         #print("Registro : ",  filtro , "No existe!!!")
         #print("IDS: ", ids)
         ident = models_prov.execute_kw(db, uid, password, 'res.partner', 'create', [{ 'name': l_prov,
                                                                                        'active': 1,}])
         return ident
        #print("id_Odoo: ", ident)
    else: return ids[0]
### Up ##############################################################
def mi_tipo_equipo(url, db, uid, password, tieq):
    import xmlrpc.client
     # Calliing methods
        
    models_tieq = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_tieq.execute_kw(db, uid, password,
                     'fincas_pma.tipo_equipo', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['name', '=', tieq], ['active','=',1]]]  #lista de python
    registros = models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'search_count', filtro)
    ids =       models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'search',       filtro, {'limit': 1})
    if registros == 0:
         #print("Registro : ",  filtro , "No existe!!!")
         #print("IDS: ", ids)
         ident = models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'create', [{ 'name': tieq,
                                                                                        'active': 1,}])
         return ident
        #print("id_Odoo: ", ident)
    else: return ids[0]
### Up ##############################################################
def mi_up(url, db, uid, password, l_up):
    import xmlrpc.client
     # Calliing methods
        
    models_up = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_up.execute_kw(db, uid, password,
                     'fincas_pma.up', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['code_up', '=', l_up], ['active','=',1]]]  #lista de python
    registros = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search_count', filtro)
    ids =       models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search',       filtro, {'limit': 1})
    if registros == 0:
         #print("Registro : ",  filtro , "No existe!!!")
         #print("IDS: ", ids)
         ident = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'create', [{ 'name': l_up,
                                                                                        'active': 1,
                                                                                        'code_up': l_up,
                                                                                        'description': l_up}])
         return ident
        #print("id_Odoo: ", ident)
    else: return ids[0]
#################################################################
# PROGRAMA PRINCIPAL - ODOO 14

#url = "http://odoradita.com:8069"
#db = "test3_CADASA_main"
url = "http://localhost:10014"
db = "t14_PU1"
username = 'soporte@alconsoft.net'
password = "2010Sistech"
max_registros = 501
###################################
import winsound
freq = 2500 # Set frequency To 2500 Hertz
dur = 1000 # Set duration To 1000 ms == 1 second
print("Beep:", winsound.Beep(freq, dur))
#Para DOS/Windows
os.system ("cls")
print("INICIANDO RUTINA DE SINCRONIZACION DE GUIAS")
common = xmlrpc.client.ServerProxy('{}/xmlrpc/2/common'.format(url))
print("common version: ")
print(common.version())

#User Identifier
uid = common.authenticate(db, username, password, {})
print("uid: ",uid)

# Calliing methods
models = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
models.execute_kw(db, uid, password,
              'res.partner', 'check_access_rights',
              ['read'], {'raise_exception': False})

##########################################################################################################
# CONSULTA DE MS-SQL EN MSSQL.ODORADTA.COM - GUIAS DE CAÑA
# SQL SERVER
cadena_conex1 = "DRIVER={SQL Server};server=mssql.odoradita.com;database=CAMPO;uid=sa;pwd=crsJVA!_02x"
conexion1 = pyodbc.connect(cadena_conex1)
cursor1 = conexion1.cursor()
# - CONSULTA MS-SQL
#SELECT Secuencia, Ano, FechaHoraCaptura, Placa, Tipo_Equipo, Tipo_Vehiculo, Contrato, Frente, Up, Proveedor, Subdiv, Fecha_Guia, Fecha_Quema, Hora_Quema
#FROM CAMPO.dbo.GUI_GUIA_CANA;
#
consulta1a = "SELECT Secuencia, Ano, convert(varchar, FechaHoraCaptura,21) as FechaHC, Placa, Tipo_Equipo, Tipo_Vehiculo, Contrato, Frente, Up, Proveedor, "
consulta1b =" Subdiv, Fecha_Guia, Fecha_Quema, Hora_Quema, Ticket, Bruto, Tara, Neto_Lbs "
consulta1 = consulta1a + consulta1b
consulta2 = "FROM CAMPO.dbo.GUI_GUIA_CANA"
consulta3 = " WHERE Dia_Zafra = 3 "
consulta4 = "ORDER BY Secuencia"
consulta = consulta1 + consulta2 + consulta3 +consulta4
print("Consulta MS-SQL: ", consulta)
cursor1.execute(consulta)
rows = cursor1.fetchall()

for row in rows:
    print(row.Secuencia, row.Ano, row[2], row.Placa, row.Tipo_Equipo, row.Frente, row.Proveedor, row.Tipo_Vehiculo, row.Fecha_Guia, row.Neto_Lbs)
    # INSERTAR REGISTROS EN TABLA purchase.order SI NO EXISTE.
    # print("Tipo de Sec." , type(row.Secuencia))
    m_secuencia = int(row.Secuencia)
    #print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
    m_fechahc = row.FechaHC
    # Tipo de Equipos
    m_tipo_equipo = mi_tipo_equipo(url, db, uid, password, row.Tipo_Equipo)
    # UP
    m_up = mi_up(url, db, uid, password, row.Up)
    # Proveedor
    m_proveedor = mi_proveedor(url, db, uid, password, row.Proveedor)
    # Bruto, Tara y Neto
    #print("Tipo Bruto: ", type(row.Bruto))
    #print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
    # Purchase Order [Encabezado]
    m_ident = 0
    filtro = [[['secuencia_guia', '=', m_secuencia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'purchase.order', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'purchase.order', 'search',       filtro, {'limit': 1})
    if registros != 0:
        m_ident = ids[0]
    if registros == 0:
        print("Registro : ",  filtro , "No existe!!!")
        #print("IDS: ", ids)
        ident = models.execute_kw(db, uid, password, 'purchase.order', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'secuencia_guia': m_secuencia,
                                                                                'ano': row.Ano,
                                                                                'zafra': 1,
                                                                                'fechahc': m_fechahc,
                                                                                'placa': row.Placa,
                                                                                'tipo_equipo': m_tipo_equipo,
                                                                                'contrato': row.Contrato,
                                                                                'frente': int(row.Frente),
                                                                                'up': m_up,
                                                                                'lote': row.Subdiv,
                                                                                'origin': row.Ticket,
                                                                                'tipo_vehiculo': row.Tipo_Vehiculo,
                                                                                'state': 'purchase',
                                                                                'fecha_guia': row.Fecha_Guia,
                                                                                'fecha_quema': row.Fecha_Quema,
                                                                                'hora_quema': row.Hora_Quema,
                                                                                'bruto': float(row.Bruto),
                                                                                'tara': float(row.Tara),
                                                                                'neto': float(row.Neto_Lbs),
                                                                                'active': True}])
        m_ident = ident
    else:
        print("Si existe registro Secuencia:", m_secuencia)
    # Purchase Order line [Detalle]
    filtro = [[['secuencia_guia', '=', m_secuencia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'purchase.order.line', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'purchase.order.line', 'search',       filtro, {'limit': 1})
    if registros == 0:
        print("Registro : ",  filtro , "No existe!!!")
        #print("IDS: ", ids)
        ident = models.execute_kw(db, uid, password, 'purchase.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'secuencia_guia': m_secuencia,
                                                                                'state': 'purchase',
                                                                                'bruto': float(row.Bruto),
                                                                                'tara': float(row.Tara),
                                                                                'neto': float(row.Neto_Lbs),
                                                                                'name': '[MP-001] CAÑA DE AZUCAR',
                                                                                'sequence':10,
                                                                                'product_qty': float(row.Neto_Lbs)*0.453592,
                                                                                'product_uom_qty': float(row.Neto_Lbs)*0.453592,
                                                                                'product_uom': 1,
                                                                                'product_id': 1,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'company_id': 1,
                                                                                'state': 'purchase',
                                                                                'qty_received_method': 'stock_moves',
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'partner_id': m_proveedor,
                                                                                'currency_id': 16,
                                                                                'active': True}])
    else:
        print("Si existe registro Secuencia:", m_secuencia)


#'zafra': row.Ano,
print("########## FIN DE RUTINA DE SINCRONIZACION DE GUIAS ############")

############ fin de programa sincronizador ##########
# -*- coding: utf-8 -*-
import pyodbc
import sys
import xmlrpc.client
import os
from datetime import datetime
from openpyxl import Workbook
import xlsxwriter
from datetime import datetime
####-PRODUCTO-#######################################################
def mi_productO(url, db, uid, password, l_defacode, l_name):
    import xmlrpc.client
    # Calliing methods
    models_prod = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_prod.execute_kw(db, uid, password,
                     'product.template', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['default_code', '=', l_defacode], ['active','=',1]]]  #lista de python
    registros = models_prod.execute_kw(db, uid, password, 'product.template', 'search_count', filtro)
    ids =       models_prod.execute_kw(db, uid, password, 'product.template', 'search',       filtro, {'limit': 1})
    if registros == 0:
         ident = models_prod.execute_kw(db, uid, password, 'product.template', 'create', [{ 'name': l_name,
                                                                                       'default_code': l_defacode,
                                                                                       'active': 1,}])
         return ident
    else: return ids[0]
####-PROVEEDOR-######################################################
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
####-EQUIPO-#########################################################
def mi_equipo(url, db, uid, password, eqid):
    import xmlrpc.client
     # Calliing methods
        
    models_eqid = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_eqid.execute_kw(db, uid, password,
                     'maintenance.equipment', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['codigo_activo', '=', eqid], ['active','=',1]]]  #lista de python
    registros = models_eqid.execute_kw(db, uid, password, 'maintenance.equipment', 'search_count', filtro)
    ids =       models_eqid.execute_kw(db, uid, password, 'maintenance.equipment', 'search',       filtro, {'limit': 1})
    if registros == 0:
         #print("Registro : ",  filtro , "No existe!!!")
         #print("IDS: ", ids)
         ident = models_eqid.execute_kw(db, uid, password, 'maintenance.equipment', 'create', [{ 'name': eqid,
                                                                                        'codigo_activo':eqid,
                                                                                        'active': 1,}])
         return ident
        #print("id_Odoo: ", ident)
    else: return ids[0]
###-TIPO EQUIPO-#####################################################
def mi_tipo_equipo(url, db, uid, password, tieq):
    import xmlrpc.client
    models_tieq = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_tieq.execute_kw(db, uid, password,
                     'fincas_pma.tipo_equipo', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['name', '=', tieq], ['active','=',1]]]  #lista de python
    registros = models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'search_count', filtro)
    ids =       models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'search',       filtro, {'limit': 1})
    if registros == 0:
         ident = models_tieq.execute_kw(db, uid, password, 'fincas_pma.tipo_equipo', 'create', [{ 'name': tieq,
                                                                                        'active': 1,}])
         return ident
    else: return ids[0]
### Up ##############################################################
def mi_up(url, db, uid, password, l_up, l_prov):
    import xmlrpc.client
     # Calliing methods
        
    models_up = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_up.execute_kw(db, uid, password,
                     'fincas_pma.up', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['code_up', '=', str(int(l_up))], ['active','=',1]]]  #lista de python
    registros = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search_count', filtro)
    ids =       models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search',       filtro, {'limit': 1})
    if registros == 0:
         #print("Registro : ",  filtro , "No existe!!!")
         #print("IDS: ", ids)
         ident = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'create', [{ 'name': str(int(l_up)),
                                                                                        'active': 1,
                                                                                        'code_up': str(int(l_up)),
                                                                                        'description': str(int(l_up)),
                                                                                        'partner_id': l_prov}])
         return ident
        #print("id_Odoo: ", ident)
    else: return ids[0]
### FRENTE ##########################################################
def mi_frente(url, db, uid, password, l_frente):
    import xmlrpc.client
    models_fren = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_fren.execute_kw(db, uid, password,
                     'fincas_pma.frentes', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['code_frente', '=', l_frente], ['active','=',1]]]  #lista de python
    registros = models_fren.execute_kw(db, uid, password, 'fincas_pma.frentes', 'search_count', filtro)
    ids =       models_fren.execute_kw(db, uid, password, 'fincas_pma.frentes', 'search',       filtro, {'limit': 1})
    if registros == 0:
         ident = models_fren.execute_kw(db, uid, password, 'fincas_pma.frentes', 'create', [{ 'name': l_frente,
                                                                                        'active': 1,
                                                                                        'code_frente': l_frente,
                                                                                        'company_id': 1,
                                                                                        'description': 'FRENTE DE PRUEBAS MAT.EXT.-DATOS FICTICIOS',
                                                                                        }])
         return ident
    else: return ids[0]
### proyecto ###################################################
def mi_proyecto(url, db, uid, password, l_up, l_lot, l_proveedor):
    import xmlrpc.client
    models_proy = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_proy.execute_kw(db, uid, password,
                     'project.project', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    lc_uplote = str(int(l_up)) + '-' + str(int(l_lot))
    lm_up = mi_up(url, db, uid, password, l_up, l_proveedor)
    filtro = [[['uplote', '=', lc_uplote], ['active','=',1]]]  #lista de python
    registros = models_proy.execute_kw(db, uid, password, 'project.project', 'search_count', filtro)
    ids =       models_proy.execute_kw(db, uid, password, 'project.project', 'search',       filtro, {'limit': 1})
    if registros == 0:
         ident = models_proy.execute_kw(db, uid, password, 'project.project', 'create', [{ 'name': l_proveedor+' '+str(int(l_up))+'-'+str(int(l_lot)),
                                                                                        'active': 1,
                                                                                        'uplote': str(int(l_up)) + '-' + str(int(l_lot)),
                                                                                        'up': lm_up,
                                                                                        'lote': str(int(l_lot)),
                                                                                        'company_id': 1,
                                                                                        'description': l_proveedor+' '+str(int(l_up))+'-'+str(int(l_lot)),
                                                                                        }])
         return ident
    else: return ids[0]
########################################################################################################################################

#url = "http://odoradita.com:8069"
#db = "test3_CADASA_main"
url = "http://localhost:8069"
db = "t14_CADASA_03"
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

con1a = "SELECT TOP (1000) [GUIA],[UP],[PROVEEDOR],[LOT],[TIPO_CANA],[FECHA_MUESTRA],[CANA_LIMPIA]"
con1b = ",[HOJAS_PORC],[CHULQUIN_PORC],[COGOLLO_PORC],[CANA_SECA_PORC],[YAGUAS_PORC]"
con1c = ",[TIERRAS_PORC],[CEPAS_PORC],[PIEDRAS_PORC],[PESO_IMPU],[CORTE_PORC],[ALCE_PORC]"
con1d = ",[IMPU_PORC],[PESO_MUESTRA],[HOJA_KG],[CHULKIN_KG],[COGOLLO_KG],[CANA_SECA_KG]"
con1e = ",[YAGUAS_KG],[TIERRA_KG],[CEPAS_KG],[PIEDRAS_KG]"
con2 =  " FROM [CAMPO].[dbo].[Agosto_Datos] "
consulta = con1a + con1b + con1c + con1d + con1e + con2
print("Consulta MS-SQL: ", consulta)
cursor1.execute(consulta)
rows = cursor1.fetchall()
for row in rows:
    print(row.GUIA, row.UP, row.LOT, row.PROVEEDOR, row.FECHA_MUESTRA, row.CANA_LIMPIA)
    m_fechahc = row.FECHA_MUESTRA.isoformat(sep=' ',timespec='seconds')
    # Tipo de Equipos
    m_equipo = mi_equipo(url, db, uid, password, '845223')
    # Proveedor
    m_proveedor = mi_proveedor(url, db, uid, password, row.PROVEEDOR)
    # UP
    m_up = mi_up(url, db, uid, password, row.UP, m_proveedor)
    # Bruto, Tara y Neto
    m_guia = str(int(row.GUIA))
    # Frente
    m_frente = mi_frente(url, db, uid, password, '10')
    # Proyecto = UPLote
    m_proyecto = mi_proyecto(url, db, uid, password, row.UP, row.LOT, row.PROVEEDOR)
    # Tipo de Caña
    if row.TIPO_CANA == 'Caña Picada Verde':
        m_tipo_cana = 'V'
    else:
        m_tipo_cana = ''
    #print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
    # Purchase Order [Encabezado]
    m_ident = 0
    filtro = [[['guia', '=', m_guia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'sample.order', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'sample.order', 'search',       filtro, {'limit': 1})
    if registros != 0:
        m_ident = ids[0]
    if registros == 0:
        print("Registro : ",  filtro , "No existe!!!")
        #print("IDS: ", ids)
        ident = models.execute_kw(db, uid, password, 'sample.order', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'zafra': 1,
                                                                                'date_order': m_fechahc,
                                                                                'equipo_id': m_equipo,
                                                                                'frente': m_frente,
                                                                                'up': m_up,
                                                                                'lote': row.LOT,
                                                                                'origin': m_guia,
                                                                                'tickete':m_guia,
                                                                                'state': 'sample',
                                                                                'diazafra':'1',
                                                                                'empleado_id':1,
                                                                                'projects_id': m_proyecto,
                                                                                'tipo_cane': m_tipo_cana
                                                                                'active': True}])
        m_ident = ident
    else:
        print("Si existe registro Secuencia:", m_guia)
    # Purchase Order line [Detalle]
    
    filtro = [[['guia', '=', m_guia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'sample.order.line', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'sample.order.line', 'search',       filtro, {'limit': 1})
    if registros == 0:
        # HOJA
        m_defacode = 'ME-001'
        m_name = '[ME-001] HOJAS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':10,
                                                                                'product_qty': float(row.HOJA_KG),
                                                                                'product_uom_qty': float(row.HOJA_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # CHULQUIN
        m_defacode = 'ME-002'
        m_name = '[ME-002] CHULQUIN'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':20,
                                                                                'product_qty': float(row.CHULKIN_KG),
                                                                                'product_uom_qty': float(row.CHULKIN_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # COGOLLOS
        m_defacode = 'ME-003'
        m_name = '[ME-003] COGOLLOS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':30,
                                                                                'product_qty': float(row.COGOLLO_KG),
                                                                                'product_uom_qty': float(row.COGOLLO_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # CAÑA SECA
        m_defacode = 'ME-004'
        m_name = '[ME-004] CAÑA SECA'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':40,
                                                                                'product_qty': float(row.CANA_SECA_KG),
                                                                                'product_uom_qty': float(row.CANA_SECA_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # YAGUAS
        m_defacode = 'ME-005'
        m_name = '[ME-005] YAGUAS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':50,
                                                                                'product_qty': float(row.YAGUAS_KG),
                                                                                'product_uom_qty': float(row.YAGUAS_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # TIERRAS
        m_defacode = 'MM-001'
        m_name = '[MM-001] TIERRAS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':60,
                                                                                'product_qty': float(row.TIERRA_KG),
                                                                                'product_uom_qty': float(row.TIERRA_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # CEPAS
        m_defacode = 'MM-002'
        m_name = '[MM-002] CEPAS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':70,
                                                                                'product_qty': float(row.CEPAS_KG),
                                                                                'product_uom_qty': float(row.CEPAS_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # PIEDRAS
        m_defacode = 'MM-003'
        m_name = '[MM-003] PIEDRAS'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':80,
                                                                                'product_qty': float(row.PIEDRAS_KG),
                                                                                'product_uom_qty': float(row.PIEDRAS_KG),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])
        # CANA LIMPIA                                                                    
        m_defacode = 'MP-002'
        m_name = '[MP-002] MUESTRA CAÑA LIMPIA'
        m_product_id = mi_productO(url, db, uid, password, m_defacode, m_name)        
        ident = models.execute_kw(db, uid, password, 'sample.order.line', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'guia': m_guia,
                                                                                'state': 'sample',
                                                                                'name': m_name,
                                                                                'sequence':90,
                                                                                'product_qty': float(row.CANA_LIMPIA),
                                                                                'product_uom_qty': float(row.CANA_LIMPIA),
                                                                                'product_uom': 1,
                                                                                'product_id': m_product_id,
                                                                                'price_unit': 0.00,
                                                                                'price_subtotal': 0.00,
                                                                                'price_total': 0.00,
                                                                                'price_tax': 0.00,
                                                                                'order_id': m_ident,
                                                                                'qty_received': 0.00,
                                                                                'qty_received_manual': 0.00,
                                                                                'active': True}])

    else:
        print("Si existe registro Secuencia:", m_guia)


print("########## FIN DE RUTINA DE SINCRONIZACION DE GUIAS ############")

############ fin de programa sincronizador ##########
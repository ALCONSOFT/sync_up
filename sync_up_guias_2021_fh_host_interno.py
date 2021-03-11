import pyodbc
import platform
from os import getenv
#import pymssql
import sys
import xmlrpc.client
import os
from datetime import datetime
from openpyxl import Workbook
import xlsxwriter
from datetime import timedelta

import time

def norma_none(lc_var1):
    if not lc_var1:
        return ''
    else:
        return lc_var1
########################################################################
def mi_proyecto(url, db, uid, password, l_up, l_lot, l_proveedor):
    import xmlrpc.client
    models_proy = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_proy.execute_kw(db, uid, password,
                     'project.project', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    lc_uplote = str(int(l_up)) + '-' + str(int(l_lot))
    lm_up = mi_up(url, db, uid, password, l_up, l_proveedor)
    lm_proveedor = mi_proveedor(url, db, uid, password, l_proveedor)
    filtro = [[['uplote', '=', lc_uplote], ['active','=',1]]]  #lista de python
    registros = models_proy.execute_kw(db, uid, password, 'project.project', 'search_count', filtro)
    idsr =      models_proy.execute_kw(db, uid, password, 'project.project', 'search_read',  filtro, {'fields':['name','uplote','partner_id'], 'limit': 1} )
    if registros == 0:
         ident = models_proy.execute_kw(db, uid, password, 'project.project', 'create', [{ 'name': l_proveedor+' '+str(int(l_up))+'-'+str(int(l_lot)),
                                                                                        'active': 1,
                                                                                        'uplote': str(int(l_up)) + '-' + str(int(l_lot)),
                                                                                        'up': lm_up,
                                                                                        'lote': str(int(l_lot)),
                                                                                        'company_id': 1,
                                                                                        'description': l_proveedor+' '+str(int(l_up))+'-'+str(int(l_lot)),
                                                                                        'partner_id': lm_proveedor,
                                                                                        }])
         return ident
    else:
        # Si el proyecto (uplote) existé; entonces hacer:
        # Verificar el Proveedor en la tabla de proyectos si existe
        if not idsr[0]['partner_id']:
            # escribirlo en la tabla de proyectos para relacionarlo al mismo
            models_proy.execute_kw(db, uid, password, 'project.project', 'write', [idsr[0]['id'] , {'partner_id': lm_proveedor}])
        return idsr[0]['id']
###################################
def mi_turno_hora(lc_hora):
    if lc_hora == '06':
        return 'Noct.'
    elif lc_hora == '07':
        return 'Diur.'
    elif lc_hora == '08':
        return 'Diur.'
    elif lc_hora == '09':
        return 'Diur.'
    elif lc_hora == '10':
        return 'Diur.'
    elif lc_hora == '11':
        return 'Diur.'
    elif lc_hora == '12':
        return 'Diur.'
    elif lc_hora == '13':
        return 'Diur.'
    elif lc_hora == '14':
        return 'Diur.'
    elif lc_hora == '15':
        return 'Diur.'
    elif lc_hora == '16':
        return 'Diur.'
    elif lc_hora == '17':
        return 'Diur.'
    elif lc_hora == '18':
        return 'Diur.'
    elif lc_hora == '19':
        return 'Noct.'
    elif lc_hora == '20':
        return 'Noct.'
    elif lc_hora == '21':
        return 'Noct.'
    elif lc_hora == '22':
        return 'Noct.'
    elif lc_hora == '23':
        return 'Noct.'
    elif lc_hora == '00':
        return 'Noct.'
    elif lc_hora == '01':
        return 'Noct.'
    elif lc_hora == '02':
        return 'Noct.'
    elif lc_hora == '03':
        return 'Noct.'
    elif lc_hora == '04':
        return 'Noct.'
    elif lc_hora == '05':
        return 'Noct.'
###################################
def mi_lote_hora(lc_hora):
    if lc_hora == '06':
        return '01:06-07'
    elif lc_hora == '07':
        return '02:07-08'
    elif lc_hora == '08':
        return '03:08-09'
    elif lc_hora == '09':
        return '04:09-10'
    elif lc_hora == '10':
        return '05:10-11'
    elif lc_hora == '11':
        return '06:11-12'
    elif lc_hora == '12':
        return '07:12-13'
    elif lc_hora == '13':
        return '08:13-14'
    elif lc_hora == '14':
        return '09:14-15'
    elif lc_hora == '15':
        return '10:15-16'
    elif lc_hora == '16':
        return '11:16-17'
    elif lc_hora == '17':
        return '12:17-18'
    elif lc_hora == '18':
        return '13:18-19'
    elif lc_hora == '19':
        return '14:19-20'
    elif lc_hora == '20':
        return '15:20-21'
    elif lc_hora == '21':
        return '16:21-22'
    elif lc_hora == '22':
        return '17:22-23'
    elif lc_hora == '23':
        return '18:23-00'
    elif lc_hora == '00':
        return '19:00-01'
    elif lc_hora == '01':
        return '20:01-02'
    elif lc_hora == '02':
        return '21:02-03'
    elif lc_hora == '03':
        return '22:03-04'
    elif lc_hora == '04':
        return '23:04-05'
    elif lc_hora == '05':
        return '24:05-06'
#################################################################
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
### Up ##########################################################
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
### Up ##########################################################
def mi_up(url, db, uid, password, l_up, l_prov):
    import xmlrpc.client
     # Calliing methods
    lm_proveedor = mi_proveedor(url, db, uid, password, l_prov)    
    models_up = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_up.execute_kw(db, uid, password,
                     'fincas_pma.up', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['code_up', '=', str(int(l_up))], ['active','=',1]]]  #lista de python
    registros = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search_count', filtro)
    idsr =      models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'search_read',  filtro, {'fields':['name','code_up','partner_id'], 'limit': 1} )
    if registros == 0:
         ident = models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'create', [{ 'name': str(int(l_up)),
                                                                                        'active': 1,
                                                                                        'code_up': str(int(l_up)),
                                                                                        'description': l_prov + '-' +str(int(l_up)),
                                                                                        'partner_id': lm_proveedor}])
         return ident
    else:
        # Si el UP existé; entonces hacer:
        # Verificar el Proveedor en la tabla de UP si existe
        if not idsr[0]['partner_id']:
            # escribirlo en la tabla de proyectos para relacionarlo al mismo
            models_up.execute_kw(db, uid, password, 'fincas_pma.up', 'write', [idsr[0]['id'] , {'partner_id': lm_proveedor}])
        return idsr[0]['id']
### Zafra #######################################################
def mi_zafra(url, db, uid, password, lc_czafra, lc_name):
    import xmlrpc.client
    models_zafra = xmlrpc.client.ServerProxy('{}/xmlrpc/2/object'.format(url))
    models_zafra.execute_kw(db, uid, password,
                     'fincas_pma.zafras', 'check_access_rights',
                     ['read'], {'raise_exception': False})
    
    filtro = [[['name', '=', lc_name], ['code_zafra', '=', lc_czafra], ['active','=',1]]]  #lista de python
    registros = models_zafra.execute_kw(db, uid, password, 'fincas_pma.zafras', 'search_count', filtro)
    ids =       models_zafra.execute_kw(db, uid, password, 'fincas_pma.zafras', 'search',       filtro, {'limit': 1})
    if registros == 0:
         ident = models_zafra.execute_kw(db, uid, password, 'fincas_pma.zafras', 'create', [{ 'name': lc_name,
                                                                                        'active': 1,
                                                                                        'code_zafra': lc_zafra,
                                                                                        'description': lc_zafra}])
         return ident
    else: return ids[0]

##################################################
def mi_sync():
    url = "http://10.11.4.213:80"
    db = "p14_CADASA_2021"
    #url = "http://localhost:80"
    #db = "p14_CADASA_2020"
    username = 'soporte@alconsoft.net'
    password = "2010Sistech"
    max_registros = 501
    ###################################
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
    # SQL SERVER POR PYODBC
    sistema = platform.system()

    if (sistema) == 'Linux':
        print("Estamos en {}".format(sistema))
        cnn = pyodbc.connect('DRIVER=FreeTDS;SERVER=10.11.4.5;PORT=1433;DATABASE=CONTROPE;UID=ecampo;PWD=Tormenta12')
        cursor1 = cnn.cursor()
    else:
        print("Estamos en {}".format(sistema))
        cadena_conex1 = "DRIVER={SQL Server};server=10.11.4.5;database=CONTROPE;uid=ecampo;pwd=Tormenta12"
        conexion1 = pyodbc.connect(cadena_conex1)
        cursor1 = conexion1.cursor()

    # - CONSULTA MS-SQL
    #print("%s %s some static text %s!"%(var_1,var_2,var_3))

    consulta1a = "SELECT Secuencia, Ano, FechaHoraCaptura, convert(varchar, FechaHoraCaptura,21) as FechaHC, Placa, Tipo_Equipo, Tipotipo_Vehiculo, Contrato, Frente, Up, Proveedor, "
    consulta1b = " Subdiv, Fecha_Guia, Fecha_Quema, Hora_Quema, Ticket, Bruto, Tara, Neto_Lbs, "
    consulta1c = " Tipo_Alce, Alce1, Alce2, Empleado_Alce1, Empleado_Alce2, Montacargas, Empleado_Montacargas, Tractor1, Tractor2, Empleado_Tractor1, Empleado_Tractor2, Nombre_Transportista, "
    consulta1d = " Num_Empleado_Transportista, Neto_Ton, Ton1, Ton2, Cha1, TCha1, Cha2, TCha2, Mula, TMula, ChaMula, TChaMula, Caja1, TCaja1, Caja2, TCaja2, "
    consulta1e = " Promedio, Dia_Zafra, Detalle, Cerrado, Eliminado, Usuario_Guia, Procesado_Contabilidad, Ticket, Hora_Entrada, Hora_Salida, CerradoTotal, "
    consulta1f = " IncentivoTL, IncentivoTI, Fecha_Tiquete, Hora_Tiquete, Usuario_Tiquete, Origen_Tiquete, Cana "
    consulta1 = "%s %s %s %s %s %s"%(consulta1a, consulta1b, consulta1c, consulta1d, consulta1e, consulta1f)
    consulta2 = "FROM dbo.GUIA"
    consulta3a = " WHERE CONVERT(VARCHAR(20), FechahoraCaptura, 120) >="
    param_fhc = "'" + (datetime.now()-timedelta(hours=2)).isoformat(sep=' ',timespec='seconds') + "'"
    consulta3b = "AND Ano=" 
    param_ano = "2021"
    #consulta3c = "AND Secuencia >"
    #consulta3 = " WHERE Ano = 2020 "
    consulta4 = "ORDER BY Secuencia"
    consulta = "%s %s %s %s %s %s %s "%(consulta1, consulta2, consulta3a, param_fhc, consulta3b, param_ano, consulta4)
    print("Consulta MS-SQL: ", consulta)

    cursor1.execute(consulta)
    rows = cursor1.fetchall()
    ###---------------------------------------------->>>>>>>>>>>>>>>>>>>>>>>
m_zafra = mi_zafra(url, db, uid, password, '2021','2020-2021')
ahora_a = datetime.now()
i = 0
for row in rows:
    i+=1
    ahora = datetime.now()
    if i > 1:
        delta = ahora - ahora_a
        ahora_a = ahora
    else:
        delta = 0.0
    print(i, ahora, delta, row.Ano, row.Dia_Zafra, row.Secuencia, row[2], row.Placa, row.Tipo_Equipo, row.Frente, row.Proveedor, row.Tipotipo_Vehiculo, row.Fecha_Guia, row.Neto_Lbs)
    # INSERTAR REGISTROS EN TABLA purchase.order SI NO EXISTE.
    # print("Tipo de Sec." , type(row.Secuencia))
    m_secuencia = int(row.Secuencia)
    #print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
    m_fechahc = row.FechaHoraCaptura.isoformat(sep=' ',timespec='seconds')
    # Tipo de Equipos
    m_tipo_equipo = mi_tipo_equipo(url, db, uid, password, row.Tipo_Equipo)
    # UP
    m_up = mi_up(url, db, uid, password, row.Up, row.Proveedor)
    # Proveedor
    m_proveedor = mi_proveedor(url, db, uid, password, row.Proveedor)
    # Proyecto
    m_proyecto = mi_proyecto(url, db, uid, password, row.Up, row.Subdiv, row.Proveedor)
    # Contrato - Equipo de Acarreo
    m_contrato = mi_equipo(url, db, uid, password, norma_none(row.Contrato))
    # Alce - Equipo de CyA
    m_alce1 = mi_equipo(url, db, uid, password, norma_none(row.Alce1))
    # Alce - Equipo de CyA
    m_alce2 = mi_equipo(url, db, uid, password, norma_none(row.Alce2))
    # Caja - Equipo Contenedor
    m_caja1 = mi_equipo(url, db, uid, password, norma_none(row.Caja1))
    # Caja - Equipo Contenedor
    m_caja2 = mi_equipo(url, db, uid, password, norma_none(row.Caja2))

    # Bruto, Tara y Neto
    #print("Tipo Bruto: ", type(row.Bruto))
    # Lote Hora
    m_hora = str(row.FechaHoraCaptura.hour).zfill(2)
    m_lote_hora = mi_lote_hora(m_hora)
    if not row.Hora_Salida:
        m_hora_Salida = "00:00"
    else:
        m_hora_Salida = row.Hora_Salida
    if not row.IncentivoTL:
        m_incentivotl = 0
    if (sistema) == 'Linux':
        m_hora_entrada = str(row.Hora_Entrada.hour) + ':' + str(row.Hora_Entrada.minute)
    else:
        m_hora_entrada = row.Hora_Entrada[0:8]
    # ANALISIS DE CAJAS Y PESO
    if row.Tipo_Equipo == 'CAMION':
        list_cajas = ['CAJA1']
        m_peso_caja1 = float(row.Neto_Ton)
        m_peso_caja2 = 0.00
    else:
        if row.Tipo_Equipo == 'TRACTOR' or row.Tipo_Equipo == 'MULA':
            list_cajas = ['CAJA1','CAJA2']
            if row.Caja2 == "0":
                m_peso_caja1 = float(row.Neto_Ton)
                m_peso_caja2 = 0.00
            else:
                m_peso_caja1 = float(row.Neto_Ton)/2
                m_peso_caja2 = float(row.Neto_Ton)/2
        else:
            list_cajas = []
            m_peso_caja1 = 0.00
            m_peso_caja2 = 0.00
    m_cant_cajas = len(list_cajas)

    # Purchase Order [Encabezado]
    m_ident = 0
    filtro = [[['secuencia_guia', '=', m_secuencia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'purchase.order', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'purchase.order', 'search',       filtro, {'limit': 1})
    if registros != 0:
        m_ident = ids[0]
    if registros == 0:
        #print("Registro : ",  filtro , "No existe!!!")
        #print("IDS: ", ids)
        ident = models.execute_kw(db, uid, password, 'purchase.order', 'create', [{ 'company_id': 1,
                                                                                'currency_id': 16,
                                                                                'partner_id': m_proveedor,
                                                                                'secuencia_guia': m_secuencia,
                                                                                'ano': row.Ano,
                                                                                'zafra': m_zafra,
                                                                                'fechahc': m_fechahc,
                                                                                'placa': row.Placa,
                                                                                'tipo_equipo': m_tipo_equipo,
                                                                                'contrato': row.Contrato,
                                                                                'frente': int(row.Frente),
                                                                                'up': m_up,
                                                                                'lote': row.Subdiv,
                                                                                'origin': row.Ticket,
                                                                                'tipo_vehiculo': row.Tipotipo_Vehiculo,
                                                                                'state': 'purchase',
                                                                                'fecha_guia': row.Fecha_Guia,
                                                                                'fecha_quema': row.Fecha_Quema,
                                                                                'hora_quema': row.Hora_Quema,
                                                                                'bruto': float(row.Bruto),
                                                                                'tara': float(row.Tara),
                                                                                'neto': float(row.Neto_Lbs),
                                                                                'date_order': m_fechahc,
                                                                                'neto_ton': float(row.Neto_Ton)/0.907185,
                                                                                'neto_tonl': float(row.Neto_Ton),
                                                                                'ton1': float(row.Ton1)/0.907185,
                                                                                'ton2': float(row.Ton2)/0.907185,
                                                                                'dia_zafra': str(row.Dia_Zafra).zfill(3),
                                                                                'lote_hora': m_lote_hora,
                                                                                'tipo_alce': row.Tipo_Alce,
                                                                                'alce1': row.Alce1,
                                                                                'alce2': row.Alce2,
                                                                                'epl_alce1': row.Empleado_Alce1,
                                                                                'epl_alce2': row.Empleado_Alce2,
                                                                                'montacargas': row.Montacargas,
                                                                                'epl_montacarga': row.Empleado_Montacargas,
                                                                                'tractor1': row.Tractor1,
                                                                                'tractor2': row.Tractor2,
                                                                                'epl_tractor1': row.Empleado_Tractor1,
                                                                                'epl_tractor2': row.Empleado_Tractor2,
                                                                                'nombre_tansportista': row.Nombre_Transportista,
                                                                                'cha1': norma_none(row.Cha1),
                                                                                'tcha1': norma_none(row.TCha1),
                                                                                'cha2': norma_none(row.Cha2),
                                                                                'tcha2': norma_none(row.TCha2),
                                                                                'mula': norma_none(row.Mula),
                                                                                'tmula': norma_none(row.TMula),
                                                                                'chamula': norma_none(row.ChaMula),
                                                                                'tchamula': norma_none(row.TChaMula),
                                                                                'caja1': norma_none(row.Caja1),
                                                                                'tcaja1': norma_none(row.TCaja1),
                                                                                'caja2': norma_none(row.Caja2),
                                                                                'tcaja2': norma_none(row.TCaja2),
                                                                                'promedio': float(row.Promedio),
                                                                                'detalle': row.Detalle,
                                                                                'cerrado': row.Cerrado,
                                                                                'eliminado': row.Eliminado,
                                                                                'usuario_guia': row.Usuario_Guia,
                                                                                'procesado_contabilidad': row.Procesado_Contabilidad,
                                                                                'hora_entrada': m_hora_entrada,
                                                                                'hora_salida': m_hora_Salida,
                                                                                'incetivo_tl': m_incentivotl,
                                                                                'incentivo_ti': row.IncentivoTI,
                                                                                'cerrado_total': row.CerradoTotal,                                                                                
                                                                                'fecha_tiquete': row.Fecha_Tiquete,
                                                                                'hora_tiquete': row.Hora_Tiquete,
                                                                                'usuario_tiquete': row.Usuario_Tiquete,
                                                                                'origen_tiquete': row.Origen_Tiquete,
                                                                                'cane': row.Cana,
                                                                                'cant_cajas': m_cant_cajas,
                                                                                'turno': mi_turno_hora(m_hora),
                                                                                'project_id': m_proyecto,
                                                                                'uplote': row.Up + '-' + row.Subdiv,
                                                                                'active': True}])
        m_ident = ident
    else:
        print("Si existe registro Secuencia:", m_secuencia)
    # Purchase Order line [Detalle]

    filtro = [[['secuencia_guia', '=', m_secuencia],['active','=',True]]]  #lista de python
    registros = models.execute_kw(db, uid, password, 'purchase.order.line', 'search_count', filtro)
    ids =       models.execute_kw(db, uid, password, 'purchase.order.line', 'search',       filtro, {'limit': 1})
    if registros == 0:
        #print("Registro : ",  filtro , "No existe!!!")
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
                                                                                'product_qty': m_peso_caja1,
                                                                                'product_uom_qty': m_peso_caja1,
                                                                                'product_uom': 1,
                                                                                'qty_received':m_peso_caja1,
                                                                                'product_id': 2,
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
                                                                                'caja': norma_none(m_caja1),
                                                                                'alce': m_alce1,
                                                                                'contrato': m_contrato,
                                                                                'project_id': m_proyecto,
                                                                                'active': True}])
    else:
        print("Si existe registro Secuencia:", m_secuencia)
# Segunda Caja: Solo tipo_equipo = 'TRACTOR' Y 'MULA' Y CAJA2 != 0
    if m_peso_caja2 != 0.00 and registros == 0:
       #print("Registro : ",  filtro , "No existe!!!")
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
                                                                                'sequence':20,
                                                                                'product_qty': m_peso_caja2,
                                                                                'product_uom_qty': m_peso_caja2,
                                                                                'product_uom': 1,
                                                                                'qty_received': m_peso_caja1,
                                                                                'product_id': 2,
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
                                                                                'caja': norma_none(m_caja2),
                                                                                'alce': m_alce2,
                                                                                'contrato': m_contrato,
                                                                                'project_id': m_proyecto,
                                                                                'active': True}])
    

#'zafra': row.Ano,
print("########## FIN DE RUTINA DE SINCRONIZACION DE GUIAS ############")

    ###----------------------------------------------<<<<<<<<<<<<<<<<<<<<<<<
#################################################################
# PROGRAMA PRINCIPAL - ODOO 14                                  #
#################################################################
# CICLO SIN FIN
i = 1
while True:
    print('Ciclo: ',i)
    mi_sync()
############ fin de programa sincronizador ##########
###################################################################################################################
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
        #m_hora_entrada = row.Hora_Entrada[0:8]
        m_hora_entrada = str(row.Hora_Entrada.hour) + ':' + str(row.Hora_Entrada.minute)
        # ANALISIS DE CAJAS Y PESO
        if row.Tipo_Equipo == 'CAMION':
            list_cajas = ['CAJA1']
            m_peso_caja1 = float(row.Neto_Lbs)
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
            print("Registro : ",  filtro , "No existe!!! - Guardandolo en Odoo; Fecha-hora:", row.Fecha_Guia,"-",m_hora_entrada)
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
                                                                                    'fechahoracaptura': m_fechahc,
                                                                                    'date_order': m_fechahc,
                                                                                    'active': True}])
                                                                                    #'date_order': row.FechaHoraCaptura,
            m_ident = ident
        else:
            print("Secuencia:", m_secuencia, "Ya existe!")
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
                                                                                    'active': True}])
        

    #'zafra': row.Ano,
    print("########## FIN DE RUTINA DE SINCRONIZACION DE GUIAS ############")
    ####2
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

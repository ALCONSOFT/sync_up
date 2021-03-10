import platform
import pyodbc

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

consulta1a = "SELECT Secuencia, Ano, FechaHoraCaptura, convert(varchar, FechaHoraCaptura,21) as FechaHC, Placa, Tipo_Equipo, Tipotipo_Vehiculo, Contrato, Frente, Up, Proveedor, "
consulta1b = " Subdiv, Fecha_Guia, Fecha_Quema, Hora_Quema, Ticket, Bruto, Tara, Neto_Lbs, "
consulta1c = " Tipo_Alce, Alce1, Alce2, Empleado_Alce1, Empleado_Alce2, Montacargas, Empleado_Montacargas, Tractor1, Tractor2, Empleado_Tractor1, Empleado_Tractor2, Nombre_Transportista, "
consulta1d = " Num_Empleado_Transportista, Neto_Ton, Ton1, Ton2, Cha1, TCha1, Cha2, TCha2, Mula, TMula, ChaMula, TChaMula, Caja1, TCaja1, Caja2, TCaja2, "
consulta1e = " Promedio, Dia_Zafra, Detalle, Cerrado, Eliminado, Usuario_Guia, Procesado_Contabilidad, Ticket, Hora_Entrada, Hora_Salida, CerradoTotal, "
consulta1f = " IncentivoTL, IncentivoTI, Fecha_Tiquete, Hora_Tiquete, Usuario_Tiquete, Origen_Tiquete, Cana "
#consulta1 = consulta1a + consulta1b + consulta1c + consulta1d + consulta1e + consulta1f
consulta1 = "%s %s %s %s %s %s"%(consulta1a, consulta1b, consulta1c, consulta1d, consulta1e, consulta1f)
consulta2 = "FROM dbo.GUIA"
consulta3a = " WHERE Dia_Zafra >="
param_dia_zafra = "2"
consulta3b = "AND Ano=" 
param_ano = "2021"
consulta3c = "AND Secuencia >"
param_sec = "2021000001"
#consulta3 = " WHERE Ano = 2020 "
consulta4 = "ORDER BY Secuencia"
consulta = "%s %s %s %s %s %s %s %s %s"%(consulta1, consulta2, consulta3a, param_dia_zafra, consulta3b, param_ano, consulta3c, param_sec, consulta4)
print("Consulta MS-SQL: ", consulta)
cursor1.execute(consulta)
rows = cursor1.fetchall()
print('Registros de Consulta: ', rows)
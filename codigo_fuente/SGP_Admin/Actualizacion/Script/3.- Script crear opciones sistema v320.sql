--> Crear nueva opción
USE [dbgtADM]
GO

IF NOT EXISTS (SELECT opc_codigo FROM a_opcsistema WHERE opc_codigo = 1121000)

begin

  INSERT INTO a_opcsistema VALUES (1121000, 'Minuta - Tabla Gramaje x Nivel', null, null)

end

--> Crear nueva opción
IF NOT EXISTS (SELECT opc_codigo FROM a_opcsistema WHERE opc_codigo = 1221000)

begin

  INSERT INTO a_opcsistema VALUES (1221000, 'Minuta - PanLed - Param. Ceco Estructura Servicio', null, null)

end

go


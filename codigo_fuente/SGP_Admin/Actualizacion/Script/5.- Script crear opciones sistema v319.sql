--> Crear nueva opción
USE [dbgtADM]
GO
IF NOT EXISTS (SELECT opc_codigo FROM a_opcsistema WHERE opc_codigo = 2190000)

begin

  INSERT INTO a_opcsistema VALUES (2190000, 'Informe - Exportar Excel So Health', null, null)

end

go


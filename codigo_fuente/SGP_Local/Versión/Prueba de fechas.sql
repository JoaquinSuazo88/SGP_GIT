/* Insertar una transacion Ingreso de usurio log_sistema*/
--insert into log_sistema (Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion)
--values (dateadd(day,-1,GETDATE()), 'ADMSGPLOCAL', 2, 'SGP', '', '', 'Crea acceso nuevo adm sgp local')

--/* 90 dias + 1 cambio de clave */
--declare @FechaInicio datetime
--declare @FechaFin datetime

--set @FechaInicio = '2021-10-05';
--set @FechaFin ='2022-01-04';
--SELECT DATEDIFF(DAY, @FechaInicio, @FechaFin);

--/* 45 dias + 1 bloqueo cuenta */
--declare @FechaInicio datetime
--declare @FechaFin datetime
--set @FechaInicio ='2022-01-04';
--set @FechaFin ='2022-02-18';
--SELECT DATEDIFF(DAY, @FechaInicio, @FechaFin);

/* 90 dias sin ocupar el sistema */
--declare @FechaInicio datetime
--declare @FechaFin datetime
--set @FechaInicio ='2022-01-04';
--set @FechaFin ='2022-04-04';
--SELECT DATEDIFF(DAY, @FechaInicio, @FechaFin);

/* Insertar una transacion Ingreso de usurio log_sistema*/
--insert into log_sistema (Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion)
--values (dateadd(day,-1,GETDATE()), 'ADMSGPLOCAL', 2, 'SGP', '', '', 'Crea acceso nuevo adm sgp local')

--select dateadd(day,-1,GETDATE())


update a_param
set par_valor = SUBSTRING('çççççîª≠«~à{|',1,20)
where par_codigo = 'csenaadm'

update a_param
set par_valor = 'çççç'
where par_codigo = 'csenaadmva'


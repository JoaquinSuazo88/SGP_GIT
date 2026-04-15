/* Crear campo IntegraAMD cliente tabla b_clientes */
USE [dbgtADM]
GO
IF NOT EXISTS ( SELECT  *
                FROM    information_schema.[columns]
                WHERE   table_name = 'b_clientes'
                        AND column_name = 'IdIntegraAMD' ) 
    BEGIN
    
	    ALTER TABLE dbo.b_clientes ADD IdIntegraAMD int default 1
    
	END
go


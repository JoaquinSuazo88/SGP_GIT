/* Crear campo Huella Carbono tabla b_ingtrediente */
USE [dbgtADM]
GO
IF NOT EXISTS ( SELECT  *
                FROM    information_schema.[columns]
                WHERE   table_name = 'b_ingrediente'
                        AND column_name = 'Huella_Carbono' ) 
    BEGIN
    
	    ALTER TABLE dbo.b_ingrediente ADD Huella_Carbono float
    
	END
go


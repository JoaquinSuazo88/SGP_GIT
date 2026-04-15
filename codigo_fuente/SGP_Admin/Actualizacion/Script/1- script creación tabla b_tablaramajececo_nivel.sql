/****** Object:  Table [dbo].[b_tablagramajececo_nivel] ******/
USE [dbgtADM]
GO
IF EXISTS ( SELECT  *
            FROM    sys.objects
            WHERE   object_id = OBJECT_ID(N'[dbo].[b_tablagramajececo_nivel]')
                    AND type IN ( N'U' ) ) 
    BEGIN

        DROP TABLE [dbo].[b_tablagramajececo_nivel]

    END

CREATE TABLE dbo.b_tablagramajececo_nivel (
    IdCeco VARCHAR(10) NOT NULL,
    IdRegimen INT NULL,
    IdTipoPlato INT NULL,
    IdIngredienteOrigen VARCHAR(10) NOT NULL,
    IdIngredienteCambio VARCHAR(10) NOT NULL,
    CantidadBruta FLOAT NULL,
	Activo varchar(1) null,
	Fecha_Creacion datetime null,
	Fecha_Modificacion datetime null,
	Usuario varchar(20)
);


-- Nivel 3: Ceco + Regimen + TipoPlato + Ingrediente
CREATE NONCLUSTERED INDEX IX_GramajeNivel_Ceco_Regimen_TipoPlato_Ingrediente
ON dbo.b_tablagramajececo_nivel (IdCeco, IdRegimen, IdTipoPlato, IdIngredienteOrigen);

-- Nivel 2: Ceco + Regimen + Ingrediente (cuando TipoPlato es NULL)
CREATE NONCLUSTERED INDEX IX_GramajeNivel_Ceco_Regimen_Ingrediente
ON dbo.b_tablagramajececo_nivel (IdCeco, IdRegimen, IdIngredienteOrigen)
WHERE IdTipoPlato IS NULL;

-- Nivel 1: Ceco + Ingrediente (cuando Regimen y TipoPlato son NULL)
CREATE NONCLUSTERED INDEX IX_GramajeNivel_Ceco_Ingrediente
ON dbo.b_tablagramajececo_nivel (IdCeco, IdIngredienteOrigen)
WHERE IdRegimen IS NULL AND IdTipoPlato IS NULL;

GO



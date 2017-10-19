IF EXISTS(SELECT 1
          FROM   INFORMATION_SCHEMA.ROUTINES
          WHERE  ROUTINE_NAME = 'TestProcedure'
                 AND SPECIFIC_SCHEMA = 'dbo')
BEGIN
  DROP PROCEDURE TestProcedure
END
GO

CREATE PROCEDURE [dbo].[TestProcedure] 
    @variable1 INTEGER  
AS
BEGIN
	CREATE table #Temp_AppConfig
	(
		config_key varchar(255),
		config_desc varchar(255) not null
	)

   insert into #Temp_AppConfig select top 10 config_key,'มกราคม' from app_config 
   select * from #Temp_AppConfig
   IF OBJECT_ID('tempdb.dbo.#Temp_AppConfig') IS NOT NULL DROP TABLE #Temp_AppConfig
END
GO
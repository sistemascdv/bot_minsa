BEGIN TRY 
	DECLARE @products TABLE
      (ItemCode            NVARCHAR(50),
        U_TiempoEntregaDias DECIMAL(8, 4)
      )
    INSERT INTO @products
    SELECT ItemCode
          ,CONVERT(DECIMAL(8, 4), isnull(U_TiempoEntregaDias,0))
    FROM   [172.16.53.64].[LABORATORIO_VIDATEC].dbo.oitm i
    WHERE  i.ItemCode LIKE 'lcl%'

    UPDATE e
    SET    e.[e_valor] = U_TiempoEntregaDias 
    FROM   @products i
           INNER JOIN labcore.dbo.estudios e
                   ON rtrim(ltrim(replace(e_descripcion,CHAR(13)+CHAR(10),''))) COLLATE SQL_Latin1_General_CP850_CI_AS = i.ItemCode
    WHERE  e.[e_valor] is null or e.[e_valor] <> U_TiempoEntregaDias
	
	select 'true',  convert(varchar,isnull(@@ROWCOUNT,0)  )
END TRY
BEGIN CATCH
    SELECT 'false', '[' + convert (varchar ,isnull(Error_number(),0) )  + ']: ' + isnull(Error_message(),'') 
END CATCH 
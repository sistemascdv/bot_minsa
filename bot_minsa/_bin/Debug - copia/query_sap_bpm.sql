SELECT     t0.updatedate                                    [Ultima act. SAP]
          ,T0.ItemCode                                      Codigo
          ,ItemName                                         Nombre
          ,U_CPT                                            CodigoCPT
          ,CONVERT(NUMERIC(8, 2), T1.[Price])               Precio
          ,T0.[validFor]
          ,T0.[frozenFor]
          ,Iif (T0.[validFor] = 'Y'
                 AND T0.[frozenFor] = 'N', 'false', 'true') Inactivo
          ,[U_Categoria]                                    [Categoria]
          ,[U_Proveedor]                                    [Proveedor]
          ,[U_Especialidad]                                 Especialidad
          ,Iif ([U_SaludFemenina] = 'Si', 'true', 'false')  [Salud Femenina]
          ,[U_TiempoDeResultados]                           [Tiempo de Resultados]
          ,Isnull([U_TiempoEntregaDias], 0)                 [Tiempo Entrega (Dias)]
          ,[U_Seccion]                                      [Seccion]
          ,[U_TipoDePrueba]                                 [Tipo de Prueba]
          ,[U_DiasDeProcesamiento]                          [Dias de Procesamiento]
          ,[U_DiasTomaDeMuestra]                            [Dias toma Muestra]
          ,[U_CondicionDelPaciente]                         [Condicion del Paciente]
          ,[U_TuboParaLaToma]                               [Tubo para la Toma]
          ,[U_TipoDeMuestra]                                [Tipo de Muestra]
          ,[U_DiasEnvioTiempoViaje]                         [Dias de Envio y Tiempo de Viaje]
          ,[U_InfoTecnicaRefrigeado]                        [Info Tecnica Refrigeado]
          ,[U_CondicionDeLaMuestra]                         [Condicion de la Muestra]
          ,[U_CantidadParaAnalisis]                         [Cantidad para Analisis]
          ,[U_Descripcion]                                  [Descripcion]
          ,'Laboratorio VIDATEC'                            [Necesidad Asociada]
FROM       oitm T0
INNER JOIN ITM1 T1
        ON T0.ItemCode = T1.ItemCode
WHERE      ( T0.ItemCode LIKE 'LCL%' )
       AND T1.[PriceList] = 1
ORDER      BY T0.ItemCode 

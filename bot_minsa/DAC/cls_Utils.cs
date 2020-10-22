using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Gph_FrameWork.Logger;
using Gph_FrameWork.Patterns.Data_Operations.Desktop;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Collections.Specialized;
using Gph_FrameWork.Emails;


namespace bot_minsa.DAC
{
    class cls_Utils
    {
        Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
        Cls_Oracle_Data_Operations oOra_Ops = new Cls_Oracle_Data_Operations();


        public cls_Utils()
        {
            oDb_Ops.evt_Exception += new Gph_FrameWork.Delegates.Cls_Delegates.dlg_Exception(oDb_Ops_evt_Exception);
            oDb_Ops.evt_Message += new Gph_FrameWork.Delegates.Cls_Delegates.dlg_Custom_Message(oDb_Ops_evt_Message);

            oOra_Ops.evt_Exception += new Gph_FrameWork.Delegates.Cls_Delegates.dlg_Exception(oOra_Ops_evt_Exception);
            oOra_Ops.evt_Message += new Gph_FrameWork.Delegates.Cls_Delegates.dlg_Custom_Message(oOra_Ops_evt_Message);
        }

        private void oOra_Ops_evt_Exception(Exception ex)
        {
            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
        }

        private void oOra_Ops_evt_Message(string pMessage)
        {
            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.DataBase, pMessage);
        }

        private void oDb_Ops_evt_Exception(Exception ex)
        {
            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
        }

        private void oDb_Ops_evt_Message(string pMessage)
        {
            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.DataBase, pMessage);
        }



        public void Update_Gensys_Employees_Info_Change_Rotative_Credit_Limit()
        {
            DataTable oEmployees_Info_Table = new DataTable();
            DataTable oTable = new DataTable();
            string vSql = "";

            try
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo la informacion basica de los colaboradores de la tabla PL_EMPLEADOS en Gensys.");

                //vSql = "SELECT "
                //     + "T1.CODEMPLEADO AS EMPLOYEE_ID,"
                //     + "T1.NOMBREEMPLEADO AS EMPLOYEE_NAME,"
                //     + "T1.NROCEDULA AS EMPLOYEE_PID,"
                //     + "T1.STATUS AS EMPLOYEE_STATUS,"
                //     + "TO_CHAR(NVL(T3.FIRST_CONTRACT_START_DATE,T1.FECINGRESO),'YYYYMMDD') AS FIRST_CONTRACT_START_DATE,"
                //     + "TO_CHAR(T1.FECINGRESO,'YYYYMMDD') AS LAST_CONTRACT_START_DATE,"
                //     + "TO_CHAR(NVL(T1.FEC_VENCE,FECSALIDA),'YYYYMMDD') AS LAST_CONTRACT_END_DATE,"
                //     + "NVL(T1.LIMITE_CREDITO_ROTATIVO,0) AS ROTATIVE_CREDIT_LIMIT,"
                //     + "NVL(T1.LIMITE_CREDITO_ROTATIVO,0) - NVL(T2.SALDO,0) AS CREDIT_BALANCE,"
                //     + "NVL(T2.SALDO,0) AS DEBIT_BALANCE "
                //     + "FROM "
                //     + "PL_EMPLEADOS T1 "
                //     + "LEFT JOIN "
                //     + "("
                //     + "SELECT "
                //     + "CODEMPLEADO,SUM(SALDO) AS SALDO "
                //     + "FROM "
                //     + "PL_VOLUNTARIA_EMPLEADO "
                //     + "WHERE "
                //     + "CODDEDUCC IN "
                //     + "("
                //     + "SELECT "
                //     + "CODDEDUCC "
                //     + "FROM "
                //     + "PL_VOLUNTARIA_MAESTRO "
                //     + "WHERE "
                //     + "NVL(GRUPODEDUC,'X') = 'V' "
                //     + "AND "
                //     + "NVL(TITULO#1_GRUPO,'X') = 'VENTAS' "
                //     + ")"
                //     + "GROUP BY "
                //     + "CODEMPLEADO "
                //     + ") T2 "
                //     + "ON "
                //     + "T1.CODEMPLEADO = T2.CODEMPLEADO "
                //     + "LEFT JOIN "
                //     + "("
                //     + "SELECT "
                //     + "CODEMPLEADO,MIN(FECINGRESO) AS FIRST_CONTRACT_START_DATE "
                //     + "FROM "
                //     + "PL_HISTORICO_CONTRATOS "
                //     + "GROUP BY "
                //     + "CODEMPLEADO "
                //     + ") T3 "
                //     + "ON "
                //     + "T1.CODEMPLEADO = T3.CODEMPLEADO ";
                vSql = "SELECT T1.NO_EMPLE AS EMPLOYEE_ID, "
                           + "T1.NOMBRE AS EMPLOYEE_NAME, "
                           + "T1.CEDULA AS EMPLOYEE_PID, "
                           + "T1.ESTADO AS EMPLOYEE_STATUS, "
                           + "TO_CHAR (T1.F_INGRESO, 'YYYYMMDD') AS FIRST_CONTRACT_START_DATE, "
                           + "TO_CHAR(T1.F_INGRESO,'YYYYMMDD') AS LAST_CONTRACT_START_DATE, "
                           + "TO_CHAR (T1.F_VENCE, 'YYYYMMDD') AS LAST_CONTRACT_END_DATE, "
                           + "NVL (T1.LIMITE_CREDITO_ROTATIVO, 0) AS ROTATIVE_CREDIT_LIMIT, "
                           + "NVL (T1.LIMITE_CREDITO_ROTATIVO, 0) - CASE WHEN  NVL (DC.SALDO, 0) <0 THEN 0 ELSE NVL (DC.SALDO, 0) END "
                           + "AS CREDIT_BALANCE, "
                           + "CASE WHEN  NVL (DC.SALDO, 0) <0 THEN 0 ELSE NVL (DC.SALDO, 0) END AS DEBIT_BALANCE "
                           + "FROM ARPLME_UNICOS T1 "
                           + "LEFT JOIN "
                           + "(  SELECT no_emple, SUM (SALDO) AS SALDO, no_cia ,cod_pla "
                           + "FROM ARPLDC "
                           + "WHERE (       NO_ACRE = (SELECT NO_ACRE FROM PL_CREDITO_ROTATIVO_DEDUC WHERE ROWNUM = 1) "
                           + "AND GRUPO IN (SELECT GRUPO FROM PL_CREDITO_ROTATIVO_DEDUC)) "
                           + "GROUP BY no_emple, no_cia ,cod_pla) DC ON DC.NO_EMPLE = T1.NO_EMPLE and dc.no_cia=t1.no_cia and dc.cod_pla=t1.cod_pla";

                oOra_Ops.Key_ConnectionString = "Gensys";
                oEmployees_Info_Table = oOra_Ops.Get_Data_in_DataTable(vSql, System.Data.CommandType.Text, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando los resultados obtenidos de la consulta a la tabla PL_EMPLEADOS en Gensys.");

                if (oEmployees_Info_Table != null)
                {
                    if (oEmployees_Info_Table.Rows.Count > 0)
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Se procedera a descargar los datos de los colaboradores a una tabla temporal.");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Cantidad de registros encontrados en la tabla PL_EMPLEADOS: [" + oEmployees_Info_Table.Rows.Count.ToString() + "]");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando la ejecucion del procedimiento almacenado [p_RotativeCredit_Delete_Tmp_Employees_Info].");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Eliminando los datos de la tabla temporal RotativeCredit_Tmp_Employees_Info.");
                        vSql = "p_RotativeCredit_Delete_Tmp_Employees_Info";
                        oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                        oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.StoredProcedure);
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizada la ejecucion del procedimiento almacenado [p_RotativeCredit_Delete_Tmp_Employees_Info].");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el proceso de volcado de datos obtenidos de los colaboradores en Gensys hacia la tabla temporal RotativeCredit_Tmp_Employees_Info.");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo la estructura de la tabla RotativeCredit_Tmp_EMployees_Info.");
                        vSql = "select * from RotativeCredit_Tmp_Employees_Info";
                        oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                        oTable = oDb_Ops.Get_Data_in_DataTable(vSql, CommandType.Text);

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Copiando los datos del objeto DataTable [oemployees_Info_Table] al objeto DataTable [oTable].");
                        foreach (DataRow oRow in oEmployees_Info_Table.Rows)
                        {
                            DataRow oNew_Row = oTable.NewRow();
                            oNew_Row["Employee_Id"] = oRow["Employee_Id"];
                            oNew_Row["Employee_Name"] = oRow["Employee_Name"];
                            oNew_Row["Employee_PID"] = oRow["Employee_PID"];
                            oNew_Row["Employee_Status"] = oRow["Employee_Status"];
                            oNew_Row["First_Contract_Start_Date"] = oRow["First_Contract_Start_Date"];
                            oNew_Row["Last_Contract_Start_Date"] = oRow["Last_Contract_Start_Date"];
                            oNew_Row["Last_Contract_End_Date"] = oRow["Last_Contract_End_Date"];
                            oNew_Row["Rotative_Credit_Limit"] = oRow["Rotative_Credit_Limit"];
                            oNew_Row["Credit_Balance"] = oRow["Credit_Balance"];
                            oNew_Row["Debit_Balance"] = oRow["Debit_Balance"];

                            oTable.Rows.Add(oNew_Row);
                        }
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizado el proceso de copiado de los datos del objeto DataTable [oemployees_Info_Table] al objeto DataTable [oTable].");


                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el volcado de los datos contenidos en el objeto DataTable [oTable] a la tabla [RotativeCredit_Tmp_Employees_Info].");
                        if (oDb_Ops.Execute_CommandBuilder(oTable, vSql))
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizado el proceso de volcado de los datos contenidos en el objeto DataTable [oTable] a la tabla [RotativeCredit_Tmp_Employees_Info].");

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando la ejecucion del procedimiento almacenado [p_RotativeCredit_Save_Employee_Info_Change_RotativeCredit_Limit].");
                            vSql = "p_RotativeCredit_Save_Employee_Info_Change_RotativeCredit_Limit";
                            oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                            oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.StoredProcedure);
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizada la ejecucion del procedimiento almacenado [p_RotativeCredit_Save_Employee_Info_Change_RotativeCredit_Limit].");
                        }
                        else
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error ocurrido al intentar realizar el volcado de los datos contenidos en el objeto DataTable [oTable] hacia la tabla temporal [RotativeCredit_Tmp_Employees_Info].");
                        }

                    }
                    else
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Verifique el proceso de obtencion de datos desde Gensys.");
                    }
                }
                else
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Verifique el proceso de obtencion de datos desde Gensys.");
                }
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }


        public void Update_Gensys_Employees_Info_Locally()
        {
            DataTable oEmployees_Info_Table = new DataTable();
            DataTable oTable = new DataTable();
            string vSql = "";

            try
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo la informacion basica de los colaboradores de la tabla PL_EMPLEADOS en Gensys.");

                //vSql = "SELECT "
                //     + "T1.CODEMPLEADO AS EMPLOYEE_ID,"
                //     + "T1.NOMBREEMPLEADO AS EMPLOYEE_NAME,"
                //     + "T1.NROCEDULA AS EMPLOYEE_PID,"
                //     + "T1.STATUS AS EMPLOYEE_STATUS,"
                //     + "TO_CHAR(NVL(T3.FIRST_CONTRACT_START_DATE,T1.FECINGRESO),'YYYYMMDD') AS FIRST_CONTRACT_START_DATE,"
                //     + "TO_CHAR(T1.FECINGRESO,'YYYYMMDD') AS LAST_CONTRACT_START_DATE,"
                //     + "TO_CHAR(NVL(T1.FEC_VENCE,FECSALIDA),'YYYYMMDD') AS LAST_CONTRACT_END_DATE,"
                //     + "NVL(T1.LIMITE_CREDITO_ROTATIVO,0) AS ROTATIVE_CREDIT_LIMIT,"
                //     + "NVL(T1.LIMITE_CREDITO_ROTATIVO,0) - NVL(T2.SALDO,0) AS CREDIT_BALANCE,"
                //     + "NVL(T2.SALDO,0) AS DEBIT_BALANCE "
                //     + "FROM "
                //     + "PL_EMPLEADOS T1 "
                //     + "LEFT JOIN "
                //     + "("
                //     + "SELECT "
                //     + "CODEMPLEADO,SUM(SALDO) AS SALDO "
                //     + "FROM "
                //     + "PL_VOLUNTARIA_EMPLEADO "
                //     + "WHERE "
                //     + "CODDEDUCC IN "
                //     + "("
                //     + "SELECT "
                //     + "CODDEDUCC "
                //     + "FROM "
                //     + "PL_VOLUNTARIA_MAESTRO "
                //     + "WHERE "
                //     + "NVL(GRUPODEDUC,'X') = 'V' "
                //     + "AND "
                //     + "NVL(TITULO#1_GRUPO,'X') = 'VENTAS' "
                //     + ")"
                //     + "GROUP BY "
                //     + "CODEMPLEADO "
                //     + ") T2 "
                //     + "ON "
                //     + "T1.CODEMPLEADO = T2.CODEMPLEADO "
                //     + "LEFT JOIN "
                //     + "("
                //     + "SELECT "
                //     + "CODEMPLEADO,MIN(FECINGRESO) AS FIRST_CONTRACT_START_DATE "
                //     + "FROM "
                //     + "PL_HISTORICO_CONTRATOS "
                //     + "GROUP BY "
                //     + "CODEMPLEADO "
                //     + ") T3 "
                //     + "ON "
                //     + "T1.CODEMPLEADO = T3.CODEMPLEADO ";
                vSql = "SELECT T1.NO_EMPLE AS EMPLOYEE_ID, "
                           + "T1.NOMBRE AS EMPLOYEE_NAME, "
                           + "T1.CEDULA AS EMPLOYEE_PID, "
                           + "T1.ESTADO AS EMPLOYEE_STATUS, "
                           + "TO_CHAR (T1.F_INGRESO, 'YYYYMMDD') AS FIRST_CONTRACT_START_DATE, "
                           + "TO_CHAR(T1.F_INGRESO,'YYYYMMDD') AS LAST_CONTRACT_START_DATE, "
                           + "TO_CHAR (T1.F_VENCE, 'YYYYMMDD') AS LAST_CONTRACT_END_DATE, "
                           + "NVL (T1.LIMITE_CREDITO_ROTATIVO, 0) AS ROTATIVE_CREDIT_LIMIT, "
                           + "NVL (T1.LIMITE_CREDITO_ROTATIVO, 0) - CASE WHEN  NVL (DC.SALDO, 0) <0 THEN 0 ELSE NVL (DC.SALDO, 0)  END "
                           + "AS CREDIT_BALANCE, "
                           + "CASE WHEN  NVL (DC.SALDO, 0) <0 THEN 0 ELSE NVL (DC.SALDO, 0)  END AS DEBIT_BALANCE "
                           + "FROM ARPLME_UNICOS T1 "
                           + "LEFT JOIN "
                           + "(  SELECT no_emple, SUM (SALDO) AS SALDO, no_cia ,cod_pla "
                           + "FROM ARPLDC "
                           + "WHERE (       NO_ACRE = (SELECT NO_ACRE FROM PL_CREDITO_ROTATIVO_DEDUC WHERE ROWNUM = 1) "
                           + "AND GRUPO IN (SELECT GRUPO FROM PL_CREDITO_ROTATIVO_DEDUC)) "
                           + "GROUP BY no_emple, no_cia ,cod_pla) DC ON DC.NO_EMPLE = T1.NO_EMPLE and dc.no_cia=t1.no_cia and dc.cod_pla=t1.cod_pla";

                oOra_Ops.Key_ConnectionString = "Gensys";
                oEmployees_Info_Table = oOra_Ops.Get_Data_in_DataTable(vSql, System.Data.CommandType.Text, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando los resultados obtenidos de la consulta a la tabla PL_EMPLEADOS en Gensys.");

                if (oEmployees_Info_Table != null)
                {
                    if (oEmployees_Info_Table.Rows.Count > 0)
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Se procedera a descargar los datos de los colaboradores a una tabla temporal.");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Cantidad de registros encontrados en la tabla PL_EMPLEADOS: [" + oEmployees_Info_Table.Rows.Count.ToString() + "]");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando la ejecucion del procedimiento almacenado [p_RotativeCredit_Delete_Tmp_Employees_Info].");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Eliminando los datos de la tabla temporal RotativeCredit_Tmp_Employees_Info.");
                        vSql = "p_RotativeCredit_Delete_Tmp_Employees_Info";
                        oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                        oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.StoredProcedure);
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizada la ejecucion del procedimiento almacenado [p_RotativeCredit_Delete_Tmp_Employees_Info].");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el proceso de volcado de datos obtenidos de los colaboradores en Gensys hacia la tabla temporal RotativeCredit_Tmp_Employees_Info.");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo la estructura de la tabla RotativeCredit_Tmp_EMployees_Info.");
                        vSql = "select * from RotativeCredit_Tmp_Employees_Info";
                        oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                        oTable = oDb_Ops.Get_Data_in_DataTable(vSql, CommandType.Text);

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Copiando los datos del objeto DataTable [oemployees_Info_Table] al objeto DataTable [oTable].");
                        foreach (DataRow oRow in oEmployees_Info_Table.Rows)
                        {
                            DataRow oNew_Row = oTable.NewRow();
                            oNew_Row["Employee_Id"] = oRow["Employee_Id"];
                            oNew_Row["Employee_Name"] = oRow["Employee_Name"];
                            oNew_Row["Employee_PID"] = oRow["Employee_PID"];
                            oNew_Row["Employee_Status"] = oRow["Employee_Status"];
                            oNew_Row["First_Contract_Start_Date"] = oRow["First_Contract_Start_Date"];
                            oNew_Row["Last_Contract_Start_Date"] = oRow["Last_Contract_Start_Date"];
                            oNew_Row["Last_Contract_End_Date"] = oRow["Last_Contract_End_Date"];
                            oNew_Row["Rotative_Credit_Limit"] = oRow["Rotative_Credit_Limit"];
                            oNew_Row["Credit_Balance"] = oRow["Credit_Balance"];
                            oNew_Row["Debit_Balance"] = oRow["Debit_Balance"];

                            oTable.Rows.Add(oNew_Row);
                        }
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizado el proceso de copiado de los datos del objeto DataTable [oemployees_Info_Table] al objeto DataTable [oTable].");


                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el volcado de los datos contenidos en el objeto DataTable [oTable] a la tabla [RotativeCredit_Tmp_Employees_Info].");
                        if (oDb_Ops.Execute_CommandBuilder(oTable, vSql))
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizado el proceso de volcado de los datos contenidos en el objeto DataTable [oTable] a la tabla [RotativeCredit_Tmp_Employees_Info].");

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando la ejecucion del procedimiento almacenado [p_RotativeCredit_Save_Employee_Info].");
                            vSql = "p_RotativeCredit_Save_Employee_Info";
                            oDb_Ops.Key_ConnectionString = "Rotative_Credit";
                            oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.StoredProcedure);
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizada la ejecucion del procedimiento almacenado [p_RotativeCredit_Save_Employee_Info].");
                        }
                        else
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error ocurrido al intentar realizar el volcado de los datos contenidos en el objeto DataTable [oTable] hacia la tabla temporal [RotativeCredit_Tmp_Employees_Info].");
                        }

                    }
                    else
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Verifique el proceso de obtencion de datos desde Gensys.");
                    }
                }
                else
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se encontraron datos en la tabla PL_EMPLEADOS en Gensys.  Verifique el proceso de obtencion de datos desde Gensys.");
                }
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }

        private bool check_Planilla_is_Run()
        {
            bool vAnswer = false;
            DataTable oTable = new DataTable();
            string vSql = "";
            try
            {
                // vSql = "SELECT * FROM PL_PAGO_EN_PROCESO WHERE NROPLAN <> 99";
                vSql = "SELECT * FROM ARPLCONTROL WHERE IND_CALCULO = 'S'";

                oOra_Ops.Key_ConnectionString = "Gensys";
                oTable = oOra_Ops.Get_Data_in_DataTable(vSql, System.Data.CommandType.Text, "");

                if (oTable != null)
                {
                    if (oTable.Rows.Count > 0)
                    {
                        vAnswer = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
            return vAnswer;
        }

        public bool check_Planilla_IsRun()
        {
            bool vPlanilla_IsRun = false;
            try
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar si no estan corriendo la planilla en Gensys.");
                vPlanilla_IsRun = check_Planilla_is_Run();

                if (vPlanilla_IsRun)
                {
                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application,"Se encontraron datos en la tabla PL_PAGO_EN_PROCESO, lo que indica que se esta corriendo la planilla en GENSYS.");
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se encontró el indicador 'S' en la tabla ARPLCONTROL, lo que indica que se esta corriendo la planilla en NAF.");
                }
                else
                {
                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application,"No se encontraron datos en la tabla PL_PAGO_EN_PROCESO, lo que indica que la planilla no ha empezado.");
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se encontró el indicador 'S' en la tabla ARPLCONTROL, lo que indica que la planilla no ha empezado.");
                }
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin de la verificacion del proceso de planilla.");
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
            return vPlanilla_IsRun;
        }

        public object get_RotativeCredit_Get_Configuration_Value(string pConfig_Category, string pConfig_SubCategory)
        {
            object vAnswer = null;
            string vSql = "";

            try
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo el valor para la Categoria de configuracion: [" + pConfig_Category + "] / Llave: [" + pConfig_SubCategory + "]");

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Llamando al procedimiento almacenado: [" + "p_RotativeCredit_Get_Configuration_Value" + "] para obtener los datos de la configuracion solicitada.");
                vSql = "p_RotativeCredit_Get_Configuration_Value";

                oDb_Ops.Key_ConnectionString = "Rotative_Credit";

                ArrayList oParams = new ArrayList();

                //add parameters
                SqlParameter oParam1 = new SqlParameter("@Config_Category", SqlDbType.NVarChar, 50);
                oParam1.Direction = ParameterDirection.Input;
                oParam1.Value = pConfig_Category;

                SqlParameter oParam2 = new SqlParameter("@Config_SubCategory", SqlDbType.NVarChar, 100);
                oParam2.Direction = ParameterDirection.Input;
                oParam2.Value = pConfig_SubCategory;

                SqlParameter oParam3 = new SqlParameter("@Config_Value", SqlDbType.Variant);
                oParam3.Direction = ParameterDirection.Output;

                oParams.Add(oParam1);
                oParams.Add(oParam2);
                oParams.Add(oParam3);

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo los datos de la configuracion.");
                oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.StoredProcedure, oParams);

                //Get Value from output parameters
                vAnswer = oDb_Ops.Parameters["@Config_Value"].Value;
                if (vAnswer != null)
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Valor obtenido para la llave [" + pConfig_SubCategory + "]: [" + vAnswer.ToString() + "]");
                }
                else
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "NO SE ECONTRARON VALORES PARA LA CONFIGURACION [" + pConfig_Category + "] / LLAVE ESPECIFICADA [" + pConfig_SubCategory + "].");
                }
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
            return vAnswer;
        }

        public void Send_Email_Log_Process(string pEmail_Subject, string pEmail_Log_File)
        {
            object vValue;
            string vRemitent = "";
            string vSMTP = "";
            string vEmail_Subject = "";
            string vFile_Attach_Path = "";
            string vEmail_Body_Template_Path = "";
            string vEmail_Templates_Path = "";
            string vEmail_Report_Template = "";

            NameValueCollection oRecipients = new NameValueCollection();
            NameValueCollection oCC_Address = new NameValueCollection();
            NameValueCollection oBCC_Address = new NameValueCollection();

            Cls_Send_Email oEmail = new Cls_Send_Email();

            try
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el proceso de envio de email de notificacion.");

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo los datos necesarios para el envio de email de notificacion.");

                vEmail_Subject = pEmail_Subject;

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [SMTP] en la tabla RotativeCredit_Configuration.");
                vValue = get_RotativeCredit_Get_Configuration_Value("System_Notifications", "SMTP");

                if (vValue != null)
                {
                    vSMTP = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_from_address] en la tabla RotativeCredit_Configuration.");
                vValue = get_RotativeCredit_Get_Configuration_Value("System_Notifications", "email_from_address");

                if (vValue != null)
                {
                    vRemitent = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_recipients_address] en la tabla RotativeCredit_Configuration.");
                vValue = get_RotativeCredit_Get_Configuration_Value("System_Notifications", "email_recipients_address");

                if (vValue != null)
                {
                    oRecipients["Recipients"] = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [Email_Templates_Path] en la tabla RotativeCredit_Configuration.");
                vValue = get_RotativeCredit_Get_Configuration_Value("System_Notifications", "email_Templates_Path");

                if (vValue != null)
                {
                    vEmail_Templates_Path = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [Email_No_Pending_Trans_To_UpLoad] en la tabla RotativeCredit_Configuration.");
                vValue = get_RotativeCredit_Get_Configuration_Value("System_Notifications", "Email_Log_Notification_Template");

                if (vValue != null)
                {
                    vEmail_Report_Template = vValue.ToString();
                }

                vFile_Attach_Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + pEmail_Log_File;
                vEmail_Body_Template_Path = vEmail_Templates_Path + vEmail_Report_Template;

                oCC_Address = null;
                oBCC_Address = null;

                oEmail.Send_Email_Log(vRemitent, oRecipients, vSMTP, vEmail_Subject, oCC_Address, oBCC_Address, vFile_Attach_Path, vEmail_Body_Template_Path, Cls_Send_Email.Email_Body_Type.Html, null);

            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }

    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using Gph_FrameWork.Logger;
using Gph_FrameWork.Patterns.Data_Operations.Desktop;
using bot_minsa.DAC;
using System.Data.SqlClient;



namespace bot_minsa.Classes
{
    class cls_Process_labcore
    {
        cls_Utils oUtils = new cls_Utils();
        Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
        Cls_Oracle_Data_Operations oOra_Ops = new Cls_Oracle_Data_Operations();

        public cls_Process_labcore()
        {
        }

        public void start_Process()
        {
            try
            {

                Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
                //DataTable oTransactions_To_Export = new DataTable();
                DataTable oTable = new DataTable();
                DataTable oTableDestiny = new DataTable();
                string vSql = "";
                string vConnection = "";
                string vTableFilter = "";
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "*********************************************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Inciando proceso de importación de datos desde SAP a Labcore.");






                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obtener tablas a importar.");
                string vTables = ConfigurationManager.AppSettings["vTables"].ToString();
                vTables = vTables.Replace(" ", "");
                var vTablesArray = vTables.Split(',');
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Tablas: [" + vTables + "].");

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Inicio de importación para cada tabla.");
                foreach (string table in vTablesArray)
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Procesando tabla: [" + table + "].");
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obtener filtro para la tabla en proceso");
                    switch (table)
                    {
                        case "oitm":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select * from oitm " + vTableFilter;
                            break;

                        case "oinv":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select * from oinv o " + vTableFilter;
                            break;

                        case "inv1":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select i.* from inv1 i inner join oinv o on o.docentry = i.docentry " + vTableFilter;
                            break;

                        case "orin":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select * from orin o " + vTableFilter;
                            break;

                        case "rin1":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select i.* from rin1 i inner join orin o on o.docentry = i.docentry " + vTableFilter;
                            break;

                        case "itm1":
                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select * from itm1 i " + vTableFilter;
                            break;
                        case "SapProveedor":

                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select CardCode, CardName,CardType,GroupCode from ocrd " + vTableFilter;
                            break;
                        case "ocrd":

                            vTableFilter = ConfigurationManager.AppSettings[table].ToString();
                            vSql = "select * from ocrd " + vTableFilter.Replace("\r\n", "");
                            break;
                        //case "":
                        //    Console.WriteLine("Case 2");
                        //    break;
                        default:
                            vSql = "select  * from " + table;
                            break;
                    }

                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Consulta a ejecutar [" + vSql + "].");
                    try
                    {




                        oDb_Ops.Key_ConnectionString = "LABORATORIO_VIDATEC";
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Ejecutando consulta en SAP.");
                        oTable = oDb_Ops.Get_Data_in_DataTable(vSql, CommandType.Text);
                        if (oTable != null)
                        {
                            if (oTable.Rows.Count > 0)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Consulta exitosa, se hallaron [" + oTable.Rows.Count.ToString() + "] registros.");
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Borrar datos en tabla destino");
                                vSql = "delete from " + table;
                                vConnection = "labcore_sap";
                                Delete_Data_in_Table(vSql, vConnection);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se procede a volcar los datos en la tabla destino");
                                string vConnBulkCopy = GetConnection(vConnection);

                                Load_Data_on_Destiny(oTable, vConnBulkCopy, table);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se completó el volcado de datos.");

                                //contar los registros en la tabla destino
                                vSql = "select count(*) from " + table;
                                oDb_Ops.Key_ConnectionString = vConnection;
                                oTableDestiny = oDb_Ops.Get_Data_in_DataTable(vSql, CommandType.Text);
                                if (oTableDestiny != null)
                                {
                                    if (oTableDestiny.Rows.Count > 0)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se insertaron [" + oTableDestiny.Rows[0][0].ToString() + "] en la tabla destino.");

                                    }
                                }
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizada la carga de datos.");
                            }
                        }
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin del proceso para la tabla actual [" + table + "].");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");

                    }
                    catch (Exception ex)
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR ***********************");
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                    }
                }
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Finalizado el grupo de tablas");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Ha finalizado el proceso de exportación.");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "*********************************************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "********* PRESIONE CUALQUIER TECLA PARA SALIR **************************");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }

        }
        public string GetConnection(string vConnectionName)
        {
            return ConfigurationManager.ConnectionStrings[vConnectionName].ConnectionString;
        }
        public void Delete_Data_in_Table(string vSql, string vConnection)
        {
            Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
            try
            {
                oDb_Ops.Key_ConnectionString = vConnection;
                oDb_Ops.Excecute_Query_Single_Result(vSql, System.Data.CommandType.Text);
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }
        public void Load_Data_on_Destiny(DataTable oTable, string vConnection, string vTableDestiny)
        {

            using (SqlConnection connection = new SqlConnection(vConnection))
            {
                connection.Open();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    //como la tabla origen y destino son identicas no es necesario mapear el datatable
                    // para cuando sean diferentes se debe habilitar las dos línas siguientes
                    //foreach (DataColumn c in oTable.Columns)
                    //    bulkCopy.ColumnMappings.Add(c.ColumnName, c.ColumnName);

                    bulkCopy.DestinationTableName = vTableDestiny;
                    bulkCopy.BulkCopyTimeout = 3600;
                    try
                    {
                        bulkCopy.WriteToServer(oTable);
                    }
                    catch (Exception ex)
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                    }
                }
            }
        }
    }
}

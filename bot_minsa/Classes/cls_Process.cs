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
using System.Collections.Specialized;
using Gph_FrameWork.Emails;
using System.IO;
using System.Net.Mail;
using bot_minsa.BPM;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace bot_minsa.Classes
{
    class cls_Process
    {
        cls_Utils oUtils = new cls_Utils();
        Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
        Cls_Oracle_Data_Operations oOra_Ops = new Cls_Oracle_Data_Operations();
        cls_bpm_entity oBPM_entity = new cls_bpm_entity();
        string vSql = "";
        string ApplicationDir = "";

        int q_insertted = 0;
        int q_updated = 0;
        int q_duplicate = 0;
        int q_processed = 0;
        bool error_global = false;

        DataTable oTableSapProducts = new DataTable();
        string sql_error;
        string sql_message;

        string cnnLABCORE = ConfigurationManager.ConnectionStrings["labcore"].ConnectionString;


        IWebDriver driver = new FirefoxDriver();
        public cls_Process()
        {
        }
        public static void log_clear()
        {
            try
            {
                //string dir = System.Environment.CurrentDirectory;
                string dir = @"C:\\bot_minsa";
                //string log = "Log.txt";
                string log = @"C:\\bot_minsa\\Log.txt";

                if (File.Exists(Path.Combine(dir, log)))
                {
                    System.IO.File.Move(Path.Combine(dir, log), Path.Combine(dir, "log_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt"));
                    File.Delete(Path.Combine(dir, log));
                }
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR BORRANDO EL LOG ANTERIOR ***********************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }

        public void start_Process()
        {
            #region MINSA

            //*************
            ///proceso para MINSA
            ///
            Cls_SQL_Data_Operations oDb_Ops = new Cls_SQL_Data_Operations();
            //DataTable oTransactions_To_Export = new DataTable();
            DataTable oTable = new DataTable();
            DataTable oTableDestiny = new DataTable();
            vSql = "";
            string vConnection = "";
            string vTableFilter = "";
            bool error_on_process = false;
            bool error_on_validation = false;
            bool recargar_pagina = false;
            int counter = 0;

            try
            {
                //inicializando el nombre del archivo de log generado por el sistema.

                log_clear();

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "*********************************************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Inciando proceso de reporte de pruebas al MINSA");
                DataTable oTableSP = new DataTable();


                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obtener datos de las pruebas a reportar (sp en labcore).");
                try
                {
                    SqlConnection cnnSP = new SqlConnection(cnnLABCORE);
                    SqlDataAdapter daSP = new SqlDataAdapter("p_reporte_minsa", cnnSP);
                    daSP.SelectCommand.CommandType = CommandType.StoredProcedure;
                    daSP.SelectCommand.CommandTimeout = 3600;
                    daSP.Fill(oTableSP);
                }
                catch (Exception ex)
                {

                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error conectando a BD Labcore / sp: p_reporte_minsa");
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                    return;
                }

                if (oTableSP != null)
                {
                    if (oTableSP.Rows.Count > 0)
                    {

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Ejecución exitosa. Cantidad de registros encontrados: [" + oTableSP.Rows.Count.ToString() + "].");

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Conectando a MINSA. ");


                        //IWebDriver driver = new FirefoxDriver();
                        //driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
                        driver.Navigate().GoToUrl("http://190.34.154.91:7050/");
                        System.Threading.Thread.Sleep(5000);
                        string minsa_user = ConfigurationManager.AppSettings["minsa_user"].ToString();
                        string minsa_pass = ConfigurationManager.AppSettings["minsa_pass"].ToString();
                        driver.FindElement(By.Id("username")).SendKeys(minsa_user);
                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                        System.Threading.Thread.Sleep(1000);
                        driver.FindElement(By.Id("password")).SendKeys(minsa_pass + Keys.Enter);
                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                        System.Threading.Thread.Sleep(1000);
                        driver.FindElement(By.Id("password")).SendKeys(Keys.Enter);
                        System.Threading.Thread.Sleep(5000);

                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);



                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Se procede a recorrer cada registro para reportarlo a MINSA. ");
                        counter = 0;
                        //se procede a hacer un recorrido de los registros  
                        foreach (DataRow oRow in oTableSP.Rows)
                        {
                            counter++;
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Nueva orden. Contador de registros: " + counter.ToString());

                            if ((counter == 1) || (recargar_pagina))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Accediendo al formulario");
                                recargar_pagina = false;
                                bool pagina_cargada = false;
                                for (int i = 1; i <= 4; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    driver.Navigate().GoToUrl("http://190.34.154.91:7050/orderentry");
                                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                                    System.Threading.Thread.Sleep(12000 + i * 1000);
                                    if (driver.FindElements(By.Id("demo_-10_value")).Count() > 0)
                                    {
                                        pagina_cargada = true;
                                        break;
                                    }
                                }
                                if (!pagina_cargada)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Después de 4 intentos no se pudo cargar la página de registro: http://190.34.154.91:7050/orderentry");
                                    return;
                                }
                            }
                            #region datos_iniciales

                            //counter++;
                            string value = "";
                            string state = "";

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application,
                                "Código estudio : [" + oRow["e_descripcion"].ToString() +
                                "], orden: [" + oRow["numero_interno"].ToString() +
                                "], Nombre: [" + oRow["cedula"].ToString() +
                                "], Nombre: [" + oRow["primer_nombre"].ToString() +
                                 "], Apellido: [" + oRow["primer_apellido"].ToString() +
                                "]");

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validación de datos. ");
                            string l_id = oRow["l_id"].ToString();
                            string tipo_documento = oRow["tipo_documento"].ToString();
                            string cedula = oRow["cedula"].ToString();
                            string primer_apellido = oRow["primer_apellido"].ToString();
                            string segundo_apellido = oRow["segundo_apellido"].ToString();
                            string primer_nombre = oRow["primer_nombre"].ToString();
                            //string segundo_nombre = oRow["segundo_nombre"].ToString();
                            string genero = oRow["genero"].ToString();
                            string genero_completo = oRow["genero_completo"].ToString();
                            string fecha_nacimiento = oRow["fecha_nacimiento"].ToString();
                            string fecha_sintomas = oRow["fecha_sintomas"].ToString();
                            string region = oRow["region"].ToString();
                            string region_completo = oRow["region_completo"].ToString();
                            string distrito = oRow["distrito"].ToString();
                            string distrito_completo = oRow["distrito_completo"].ToString();
                            string corregimiento = oRow["corregimiento"].ToString();
                            string corregimiento_completo = oRow["corregimiento_completo"].ToString();
                            string direccion = oRow["direccion"].ToString();
                            string persona_contacto = oRow["persona_contacto"].ToString();
                            string telefono_contacto = oRow["telefono_contacto"].ToString();
                            string correo = oRow["correo"].ToString();
                            string telefono = oRow["telefono"].ToString().Replace("+507", "");
                            string tipo_de_orden = oRow["tipo_de_orden"].ToString();
                            string tipo_de_orden_completo = oRow["tipo_de_orden_completo"].ToString();
                            string numero_interno = oRow["numero_interno"].ToString();
                            //string procedencia_muestra = oRow["procedencia_muestra"].ToString();
                            string fecha_de_toma = oRow["fecha_de_toma"].ToString();
                            string tipo_de_prueba = oRow["tipo_de_prueba"].ToString();
                            string tipo_de_prueba_completo = oRow["tipo_de_prueba_completo"].ToString();
                            string resultado_minsa = oRow["resultado_minsa"].ToString();
                            string resultado_minsa_completo = oRow["resultado_minsa_completo"].ToString();
                            string resultado_valor = oRow["resultado_valor"].ToString();
                            //string resultado_igg = oRow["resultado_igg"].ToString();
                            //string resultado_igm = oRow["resultado_igm"].ToString();
                            string fecha_resultado = oRow["fecha_resultado"].ToString();
                            string tipo_de_paciente = oRow["tipo_de_paciente"].ToString();
                            string tipo_de_paciente_completo = oRow["tipo_de_paciente_completo"].ToString();
                            string tipo_de_muestra = oRow["tipo_de_muestra"].ToString();
                            string tipo_de_muestra_completo = oRow["tipo_de_muestra_completo"].ToString();


                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Numero de orden: " + numero_interno);
                            #endregion

                            #region validacion_de_datos
                            error_on_validation = false;
                            if (String.IsNullOrEmpty(primer_nombre))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "primer_nombre en blanco. ");
                                error_on_validation = true;
                            }
                            if (String.IsNullOrEmpty(primer_apellido))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "primer_apellido en blanco. ");
                                error_on_validation = true;
                            }
                            if (String.IsNullOrEmpty(genero))
                            {
                                //buscar si 
                                string genero_minsa = driver.FindElement(By.Id("demo_-104_value")).GetAttribute("value").Trim();
                                if (String.IsNullOrEmpty(genero_minsa))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "genero en blanco. ");
                                    error_on_validation = true;
                                }
                                else
                                {
                                    genero = genero_minsa.Replace(".", ". ");
                                }
                            }

                            if (String.IsNullOrEmpty(fecha_nacimiento))
                            {
                                string fecha_nacimiento_minsa = driver.FindElement(By.Id("genero")).GetAttribute("value").Trim();
                                if (String.IsNullOrEmpty(fecha_nacimiento_minsa))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_nacimiento en blanco. ");
                                    error_on_validation = true;
                                }
                                else
                                {
                                    fecha_nacimiento = fecha_nacimiento_minsa;
                                }
                            }
                            if (String.IsNullOrEmpty(region))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "region en blanco. ");
                                error_on_validation = true;
                            }
                            if (String.IsNullOrEmpty(distrito))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "distrito en blanco. Se toma 95. SIN DEFINIR");
                                //error_on_validation = true;
                                distrito = "95."; //95. SIN DEFINIR
                                distrito_completo = "95. SIN DEFINIR";
                            }
                            if (String.IsNullOrEmpty(corregimiento))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "corregimiento en blanco. Se toma 873. SIN DEFINIR");
                                //error_on_validation = true;
                                corregimiento = "873."; //873. SIN DEFINIR
                                corregimiento = "873. SIN DEFINIR";
                            }
                            if (resultado_valor == "1") //para resultados positivos es obligatorio un telefono
                            {
                                if (String.IsNullOrEmpty(telefono))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "telefono en blanco. ");
                                    error_on_validation = true;
                                }
                            }

                            if (error_on_validation)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Hay errores en validacion para este estudio, no se puede reportar.  ");

                                //try to update 'laboratorio' to manual report because error_on_validation==true
                                update_labcore_order(l_id, "3");
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Pasando a la siguiente orden. ");
                                continue; //next for 
                            }
                            #endregion


                            try
                            {



                                #region activar_formulario

                                bool paciente_existe = false;
                                //System.Threading.Thread.Sleep(5000);
                                IWebElement button_cancelar = null;
                                if (TryFindElement(By.XPath("//button[contains(text(), 'Cancelar')]"), out button_cancelar))
                                {
                                    if (button_cancelar.Displayed && button_cancelar.Enabled)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón cancelar activo. ");
                                        try
                                        {
                                            button_cancelar.Click();
                                        }
                                        catch (Exception)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón cancelar no se pudo cliquear. ");
                                        }
                                    }
                                }

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando boton 'Nuevo'. ");
                                var button_nuevo = driver.FindElement(By.XPath("//*[@ng-click='vm.eventNew()']"));
                                bool next_foreach = false;

                                for (int i = 1; i <= 4; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 2000);
                                    if ((button_nuevo.Displayed && button_nuevo.Enabled))
                                    {
                                        try
                                        {
                                            button_nuevo.Click();
                                            next_foreach = false;
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en 'Nuevo'.");
                                            break;
                                        }
                                        catch (Exception)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón 'Nuevo' no se pudo cliquear. ");
                                        }
                                    }
                                    else { next_foreach = true; }
                                }
                                if (next_foreach)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Despues de 4 intentos no se pudo cliquear botón 'Nuevo'.");
                                    continue;
                                }
                                #endregion

                                #region Validacion_Paciente

                                System.Threading.Thread.Sleep(1000);
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_documento.");
                                //Tipo documento *	demo_-10_value	tipo_documento
                                driver.FindElement(By.Id("demo_-10_value")).Click();
                                driver.FindElement(By.Id("demo_-10_value")).SendKeys(tipo_documento + Keys.Enter);
                                System.Threading.Thread.Sleep(1500);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en cedula.");
                                //Cédula *	demo_-100	cedula
                                driver.FindElement(By.Id("demo_-100")).Click();
                                driver.FindElement(By.Id("demo_-100")).SendKeys(cedula + Keys.Enter);
                                System.Threading.Thread.Sleep(1500);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si existe el paciente.");
                                //validar si existe el paciente
                                var element = WaitForElement(1, By.Id("demo_-103"), driver);
                                value = element.GetAttribute("value").Trim();
                                //value = driver.FindElement(By.Id("demo_-103")).GetAttribute("value").Trim();

                                if (!String.IsNullOrEmpty(value))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Paciente existe.");
                                    string comentario_de_la_orden = "";
                                    paciente_existe = true;

                                    //evauar nombre 1
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si primer nombre de labcore es diferente de primer nombre de minsa.");
                                    //Primer nombre *	demo_-103	primer_nombre
                                    string[] primer_nombre_values = driver.FindElement(By.Id("demo_-103")).GetAttribute("value").Trim().Split(' ');
                                    string primer_nombre_minsa = primer_nombre_values[0];

                                    if (primer_nombre != primer_nombre_minsa)
                                    {
                                        comentario_de_la_orden = "PACIENTE: " + primer_nombre + " " + primer_apellido + "." + Keys.Enter;
                                    }

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si fecha nacimiento de labcore es diferente de fecha nacimiento de minsa.");
                                    //evaluar fecha nacimiento
                                    string fecha_nacimiento_value = driver.FindElement(By.Id("demo_-105")).GetAttribute("value").Trim();
                                    if (fecha_nacimiento != fecha_nacimiento_value)
                                    {
                                        comentario_de_la_orden += "FECHA NACIMIENTO: " + fecha_nacimiento + "." + Keys.Enter;
                                    }
                                    if (!String.IsNullOrEmpty(comentario_de_la_orden))
                                    {


                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en comentario_de_la_orden: " + comentario_de_la_orden);
                                        IWebElement commentario_box = null;
                                        bool commentario = TryFindElement(By.XPath("//*[@ng-model='commentorder.notes[0].commentArray.content']"), out commentario_box);
                                        if (commentario)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ng-model='commentorder.notes[0].commentArray.content' existe.");
                                        }
                                        else
                                        {
                                            commentario = TryFindElement(By.Id("ui-tinymce-6"), out commentario_box);
                                            if (commentario)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ui-tinymce-6' existe.");
                                            }
                                            else
                                            {
                                                commentario = TryFindElement(By.Id("ui-tinymce-12"), out commentario_box);
                                                if (commentario)
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ui-tinymce-12' existe.");
                                                }
                                            }

                                        }
                                        if (commentario)
                                        {
                                            try
                                            {
                                                //en caso de no coincidir el nombre añadir nota
                                                //Comentario de la orden	ui-tinymce-6	comentario_de_la_orden
                                                commentario_box.Click();
                                                commentario_box.SendKeys(comentario_de_la_orden + Keys.Enter);
                                                System.Threading.Thread.Sleep(1000);
                                            }
                                            catch (Exception)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error al añadir el comentario.");
                                            }
                                        }
                                        else
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Comentario textbox no hallado.");
                                        }
                                    }
                                }
                                #endregion



                                #region ingreso_de_datos

                                if (!paciente_existe)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en primer_apellido.");
                                    //Primer apellido *	demo_-101	primer_apellido
                                    var element_primer_apellido = WaitForElement(2, By.Id("demo_-101"), driver);
                                    System.Threading.Thread.Sleep(600);
                                    element_primer_apellido.Click();
                                    element_primer_apellido.SendKeys(primer_apellido + Keys.Enter);
                                    System.Threading.Thread.Sleep(1000);

                                    //Primer apellido *	demo_-101	primer_apellido
                                    //driver.FindElement(By.Id("demo_-101")).Click();
                                    //driver.FindElement(By.Id("demo_-101")).SendKeys(primer_apellido + Keys.Enter);
                                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0.3);
                                    //System.Threading.Thread.Sleep(1000);

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en segundo_apellido.");
                                    //Segundo apellido	demo_-102	segundo_apellido
                                    driver.FindElement(By.Id("demo_-102")).Click();
                                    driver.FindElement(By.Id("demo_-102")).SendKeys(segundo_apellido + Keys.Enter);
                                    System.Threading.Thread.Sleep(600);

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en primer_nombre.");
                                    //Primer nombre *	demo_-103	primer_nombre
                                    driver.FindElement(By.Id("demo_-103")).Click();
                                    driver.FindElement(By.Id("demo_-103")).SendKeys(primer_nombre + Keys.Enter);
                                    System.Threading.Thread.Sleep(600);

                                    ////Segundo nombre	demo_-103	segundo_nombre
                                    //driver.FindElement(By.Id("demo_-103")).SendKeys(segundo_nombre + Keys.Enter); //no lo usamos
                                    //System.Threading.Thread.Sleep(1000);

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en genero.");
                                    //Género *	demo_-104_value	genero
                                    driver.FindElement(By.Id("demo_-104_value")).Click();
                                    driver.FindElement(By.Id("demo_-104_value")).SendKeys(genero + Keys.Enter);
                                    System.Threading.Thread.Sleep(600);
                                }//if (!paciente_existe)

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_nacimiento.");
                                //Fecha nacimiento *	demo_-105	fecha_nacimiento
                                driver.FindElement(By.Id("demo_-105")).Click();
                                driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_-105")).SendKeys(fecha_nacimiento);
                                System.Threading.Thread.Sleep(600);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_sintomas.");
                                //FECHA DE SÍNTOMAS *	demo_7	fecha_sintomas
                                driver.FindElement(By.Id("demo_7")).Click();
                                driver.FindElement(By.Id("demo_7")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_7")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_7")).SendKeys(fecha_sintomas);
                                System.Threading.Thread.Sleep(600);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en region.");
                                //REGIÓN *	demo_1_value	region
                                driver.FindElement(By.Id("demo_1_value")).Clear();
                                driver.FindElement(By.Id("demo_1_value")).Click();
                                driver.FindElement(By.Id("demo_1_value")).SendKeys(region + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en distrito.");
                                //DISTRITO *	demo_2_value	distrito
                                driver.FindElement(By.Id("demo_2_value")).Clear();
                                driver.FindElement(By.Id("demo_2_value")).Click();
                                driver.FindElement(By.Id("demo_2_value")).SendKeys(distrito + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en corregimiento.");
                                //CORREGIMIENTO *	demo_3_value	corregimiento
                                driver.FindElement(By.Id("demo_3_value")).Clear();
                                driver.FindElement(By.Id("demo_3_value")).Click();
                                driver.FindElement(By.Id("demo_3_value")).SendKeys(corregimiento + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                if (!String.IsNullOrEmpty(direccion))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en direccion.");
                                    //Dirección	demo_-112	direccion
                                    driver.FindElement(By.Id("demo_-112")).Clear();
                                    driver.FindElement(By.Id("demo_-112")).SendKeys(direccion + Keys.Enter);
                                    System.Threading.Thread.Sleep(1000);
                                }
                                //PERSONA CONTACTO	demo_5	persona_contacto
                                //TELÉFONO CONTACTO	demo_6	telefono_contacto

                                //if (!String.IsNullOrEmpty(correo))
                                //{
                                //    string[] correo_values = correo.Trim().Split(',');
                                //    string primer_correo = correo_values[0];
                                //    //Correo	demo_-106	correo
                                //    driver.FindElement(By.Id("demo_-106")).Clear();
                                //    driver.FindElement(By.Id("demo_-106")).SendKeys(primer_correo + Keys.Enter);
                                //    System.Threading.Thread.Sleep(1000);
                                //}

                                if (!String.IsNullOrEmpty(telefono))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en telefono.");
                                    //Teléfono	demo_-111	telefono
                                    driver.FindElement(By.Id("demo_-111")).Clear();
                                    driver.FindElement(By.Id("demo_-111")).SendKeys(telefono + Keys.Enter);
                                    System.Threading.Thread.Sleep(600);
                                }

                                ////Tipo de orden *	demo_-4_value	tipo_de_orden
                                //driver.FindElement(By.Id("demo_-4_value")).Clear();
                                //driver.FindElement(By.Id("demo_-4_value")).SendKeys(tipo_de_orden_completo + Keys.Tab);
                                ////System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en numero_interno.");
                                //NÚMERO INTERNO *	demo_8	numero_interno
                                driver.FindElement(By.Id("demo_8")).Click();
                                driver.FindElement(By.Id("demo_8")).SendKeys(numero_interno + Keys.Enter);
                                System.Threading.Thread.Sleep(600);

                                ////PROCEDENCIA MUESTRA	demo_9_value	procedencia_muestra 
                                //driver.FindElement(By.Id("demo_9_value")).SendKeys(oRow["procedencia_muestra"].ToString() + Keys.Enter);
                                ////System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_de_toma.");
                                //FECHA DE TOMA	demo_10	fecha_de_toma 
                                driver.FindElement(By.Id("demo_10")).Click();
                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_10")).SendKeys(fecha_de_toma);
                                System.Threading.Thread.Sleep(600);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_prueba.");
                                //TIPO DE PRUEBA *	demo_12_value	tipo_de_prueba 
                                driver.FindElement(By.Id("demo_12_value")).Click();
                                driver.FindElement(By.Id("demo_12_value")).SendKeys(tipo_de_prueba + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en resultado_minsa.");
                                //RESULTADO *	demo_11_value	resultado_minsa 
                                driver.FindElement(By.Id("demo_11_value")).Click();
                                driver.FindElement(By.Id("demo_11_value")).SendKeys(resultado_minsa + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                //RESULTADO IGG *	demo_16_value	resultado_igg
                                //RESULTADO IGM *	demo_17_value	resultado_igm

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_resultado.");
                                //FECHA RESULTADO *	demo_13	fecha_resultado
                                driver.FindElement(By.Id("demo_13")).Click();
                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                driver.FindElement(By.Id("demo_13")).SendKeys(fecha_resultado);
                                System.Threading.Thread.Sleep(600);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_paciente.");
                                //TIPO DE PACIENTE	demo_14_value	tipo_de_paciente
                                driver.FindElement(By.Id("demo_14_value")).Click();
                                driver.FindElement(By.Id("demo_14_value")).SendKeys(tipo_de_paciente + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_muestra.");
                                //TIPO DE MUESTRA	demo_15_value	tipo_de_muestra
                                driver.FindElement(By.Id("demo_15_value")).Click();
                                driver.FindElement(By.Id("demo_15_value")).SendKeys(tipo_de_muestra_completo + Keys.Enter);
                                System.Threading.Thread.Sleep(1000);

                                #endregion

                                #region grabar_datos



                                //grabar registro con ALT + S
                                var button_guardar = driver.FindElement(By.XPath("//*[@ng-click='vm.eventSave()']"));
                                try
                                {
                                    button_guardar.Click();
                                    //validar que registro se guardado
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando si boton 'Nuevo' esta desactivado y boton guardar está activo. ");
                                    bool error_al_guardar = false;
                                    for (int i = 1; i <= 4; i++)
                                    {
                                        System.Threading.Thread.Sleep(4000 + i * 1000);
                                        //evaluar si boton nuevo esta disabled
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Intento: " + i.ToString());
                                        IWebElement button_nuevo_verificacion = null;// driver.FindElement(By.XPath("//*[@ng-click='vm.eventNew()']"));
                                        if (TryFindElement(By.XPath("//*[@ng-click='vm.eventNew()']"), out button_nuevo_verificacion))
                                        {
                                            if ((!button_nuevo_verificacion.Enabled) && button_guardar.Enabled)
                                            {
                                                error_al_guardar = true;
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "El boton 'Nuevo' esta desactivado.");
                                                //no se ha terminado de guardar.
                                                //esperar
                                            }
                                            else
                                            {
                                                error_al_guardar = false;
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "El boton 'Nuevo' esta activado");
                                                break;
                                            }
                                        }
                                    }

                                    if (error_al_guardar)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "No he habilitó el botón nuevo, se colocará este registro para revisión manual");
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Puede tratarse de registro duplicado.");
                                        //update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual

                                        //intertar click en boton deshacer
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Buscar botón 'deshacer' y dar click.");
                                        var boton_deshacer = driver.FindElement(By.XPath("//*[@ng-click='vm.eventUndo()']"));
                                        if ((boton_deshacer.Displayed && boton_deshacer.Enabled))
                                        {
                                            try
                                            {
                                                boton_deshacer.Click();
                                                System.Threading.Thread.Sleep(5000);
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en 'Deshacer'.");
                                                update_labcore_order(l_id, "3");
                                            }
                                            catch (Exception)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón 'Deshacer' no se pudo cliquear. ");
                                                recargar_pagina = true;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //no hubo error se registra como enviado al minsa.
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Datos registrados. ");
                                        update_labcore_order(l_id, "1");
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error en click guardar = '" + oRow["l_id"].ToString() + "'");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                                    error_on_process = true;

                                }
                                //System.Threading.Thread.Sleep(10000);
                                if (error_on_process)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "error_on_process == true.");
                                    if (!String.IsNullOrEmpty(l_id))
                                    {
                                        update_labcore_order(l_id, "3");
                                    }
                                }

                                #endregion
                            }
                            catch (Exception ex)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error durante la inserción de datos en la página.");
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());

                                break; //end for
                            }//end try 

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin de la orden actual. ");

                        }//end foreach (DataRow oRow in oTableSP.Rows)
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin de las órdenes. ");
                    }
                    else
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No hubo registros insertados.");
                    } // end if (oTableSP.Rows.Count > 0)
                } //end if (oTableSP != null)

                //en caso de error enviar correo de advertencia
                //if (error_global)


            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR ***********************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                error_global = true;
            }

            if (error_global || error_on_validation || error_on_process)
            {
                Send_Email(error_global, "Alerta: Error en proceso");
            }

            #endregion


        }
        //end public void start_Process()

        private static void SqlExecuteNonQuery(string queryString, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
        }//end private static void SqlExecuteNonQuery(string queryString, string connectionString)

        public string GetConnection(string vConnectionName)
        {
            return ConfigurationManager.ConnectionStrings[vConnectionName].ConnectionString;
        }


        public void email_send()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["Smtp_Server.Host"]);
                mail.From = new MailAddress(ConfigurationManager.AppSettings["e_mail.From"]);
                mail.To.Add(ConfigurationManager.AppSettings["buzon_alertas"]);
                mail.Subject = "Proceso de actualización SAP - BPM.";
                mail.Body = "Proceso de actualización SAP - BPM. Se anexa log del proceso.";


                string dir = System.Environment.CurrentDirectory;
                string file = "Log.txt";

                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(Path.Combine(dir, file));
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["Smtp_Server.User"], ConfigurationManager.AppSettings["Smtp_Server.Pass"]);
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }

        public void Send_Email_Log()
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
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el proceso de envio de email de notificacion.");

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo los datos necesarios para el envio de email de notificacion.");

                vEmail_Subject = "Proceso de reporte MINSA.";

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [SMTP].");
                vValue = ConfigurationManager.AppSettings["Smtp_Server.Host"];

                if (vValue != null)
                {
                    vSMTP = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_from_address].");
                vValue = ConfigurationManager.AppSettings["e_mail.From"];

                if (vValue != null)
                {
                    vRemitent = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_recipients_address].");
                vValue = ConfigurationManager.AppSettings["buzon_alertas"];

                if (vValue != null)
                {
                    oRecipients["Recipients"] = vValue.ToString();
                }

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_log_emplate].");
                string dir = System.Environment.CurrentDirectory;
                //string file = "email_log_emplate.txt";
                string file = @"C:\\bot_minsa\\email_log_emplate.txt";
                vValue = Path.Combine(dir, file);

                if (vValue != null)
                {
                    vEmail_Templates_Path = vValue.ToString();
                }

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [Email_No_Pending_Trans_To_UpLoad] en la tabla RotativeCredit_Configuration.");
                vValue = "";// oUtils.get_RotativeCredit_Get_Configuration_Value("System_Notifications", "Email_Log_Notification_Template");

                if (vValue != null)
                {
                    vEmail_Report_Template = vValue.ToString();
                }

                //vFile_Attach_Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Log.txt";
                vFile_Attach_Path = @"C:\\bot_minsa\\Log.txt";
                vEmail_Body_Template_Path = vEmail_Templates_Path + vEmail_Report_Template;

                oCC_Address = null;
                oBCC_Address = null;

                oEmail.Send_Email_Log(vRemitent, oRecipients, vSMTP, vEmail_Subject, oCC_Address, oBCC_Address, vFile_Attach_Path, vEmail_Body_Template_Path, Cls_Send_Email.Email_Body_Type.Html, null);

            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }


            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "********* PRESIONE CUALQUIER TECLA PARA SALIR **************************");
            Console.ReadKey();

        }


        public void Send_Email(bool error, string vEmail_Subject)
        {
            object vValue;
            string vRemitent = "";
            string vSMTP = "";
            //string vEmail_Subject = "";
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
                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Iniciando el proceso de envio de email de notificacion.");

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obteniendo los datos necesarios para el envio de email de notificacion.");

                //vEmail_Subject = "Proceso de actualización SAP - BPM.";

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [SMTP].");
                vValue = ConfigurationManager.AppSettings["Smtp_Server.Host"];

                if (vValue != null)
                {
                    vSMTP = vValue.ToString();
                }

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_from_address].");
                vValue = ConfigurationManager.AppSettings["e_mail.From"];

                if (vValue != null)
                {
                    vRemitent = vValue.ToString();
                }

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_recipients_address].");
                vValue = ConfigurationManager.AppSettings["buzon_alertas"];

                if (vValue != null)
                {
                    oRecipients["Recipients"] = vValue.ToString();
                }

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [email_log_emplate].");
                string dir = System.Environment.CurrentDirectory;
                //string file = "email_log_emplate.txt";
                string file = "";
                if (error)
                {
                    file = @"C:\\bot_minsa\\email_error_template.txt";
                }
                else
                {
                    file = @"C:\\bot_minsa\\email_log_emplate.txt";
                }
                vValue = Path.Combine(dir, file);

                if (vValue != null)
                {
                    vEmail_Templates_Path = vValue.ToString();
                }

                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando el valor para la llave [Email_No_Pending_Trans_To_UpLoad] en la tabla RotativeCredit_Configuration.");
                vValue = "";// oUtils.get_RotativeCredit_Get_Configuration_Value("System_Notifications", "Email_Log_Notification_Template");

                if (vValue != null)
                {
                    vEmail_Report_Template = vValue.ToString();
                }

                ////vFile_Attach_Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Log.txt";
                //vFile_Attach_Path = @"C:\\bot_minsa\\Log.txt";
                //vEmail_Body_Template_Path = vEmail_Templates_Path + vEmail_Report_Template;

                oCC_Address = null;
                oBCC_Address = null;

                oEmail.Send_Email(vRemitent, oRecipients, vSMTP, vEmail_Subject, oCC_Address, oBCC_Address, null, vEmail_Templates_Path, Cls_Send_Email.Email_Body_Type.Text, null);//(vRemitent, oRecipients, vSMTP, vEmail_Subject, oCC_Address, oBCC_Address, vFile_Attach_Path, vEmail_Body_Template_Path, Cls_Send_Email.Email_Body_Type.Html, null);

            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }


        public bool TryFindElement(By by, out IWebElement element)
        {
            element = null;
            try
            {
                element = driver.FindElement(by);
            }
            catch (NoSuchElementException ex)
            {
                return false;
            }
            return true;
        }

        public IWebElement WaitForElement(int seconds, By By, IWebDriver driver)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));
            return wait.Until(drv => drv.FindElement(By));
        }

        public bool update_labcore_order(string l_id, string _minsa_enviado)
        {
            try  //update labcore.dbo.laboratorio.l_minsa_enviado = 3 // 3 = validation error
            {
                vSql = "UPDATE "
                        + "LABORATORIOS "
                        + "SET "
                        + "l_minsa_enviado = '" +
                        _minsa_enviado + "' "
                        + "WHERE "
                        + "l_id = '" + l_id
                        + "'";
                SqlExecuteNonQuery(vSql, cnnLABCORE);
                return true;

            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error actualizando el registro: LABCORE.dbo.LABORATORIOS l_id = '" + l_id + "'");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Consulta: " + _minsa_enviado);
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                return false;

            }
        }

    }




}


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
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.Events;
//using OpenQA.Selenium.Support.PageObjects;

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
        //FirefoxDriver driver = new FirefoxDriver();

        public static IWebDriver driver = new FirefoxDriver();

        //// Wrapping parent driver             
        //EventFiringWebDriver eventFiringWebDriver = new EventFiringWebDriver(driver);

        //// Attaching events             
        //// Attaching click events      
        //eventFiringWebDriver
        //eventFiringWebDriver.ElementClicking += EventFiringWebDriver_ElementClicking;

        ////Click events
        //// Attaching click events             
        //eventFiringWebDriver.ElementClicking += EventFiringWebDriver_ElementClicking; 
        //eventFiringWebDriver.ElementClicked += EventFiringWebDriver_ElementClicked; 



        public cls_Process()
        {
        }

        public static void log_clear()
        {
            try
            {
                string dir = System.Environment.CurrentDirectory;
                //string dir = @"C:\\bot_minsa";
                string log = "Log.txt";
                //string log = @"C:\\bot_minsa\\Log.txt";

                if (File.Exists(Path.Combine(dir, log)))
                {
                    System.IO.File.Move(Path.Combine(dir, log), Path.Combine(dir, "log_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".txt"));
                    File.Delete(Path.Combine(dir, log));
                }
            }
            catch (Exception ex)
            {
                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR BORRANDO EL LOG ANTERIOR ***********************");
                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
        }
        //private void EventFiringWebDriver_ElementClicking(object sender, WebElementEventArgs e)
        //{
        //    Console.WriteLine("Element clicking");
        //    Console.WriteLine(e);
        //    Console.WriteLine(sender);

        //}
        //private void EventFiringWebDriver_FindingElement(object sender, FindElementEventArgs e)
        //{
        //    Console.WriteLine("Finding element");
        //}

        //private void EventFiringWebDriver_FindElementCompleted(object sender, FindElementEventArgs e)
        //{
        //    Console.WriteLine("Finding element completed");
        //}

        public void start_Process()
        {

            // Wrapping parent driver             
            //EventFiringWebDriver eventFiringWebDriver = new EventFiringWebDriver(driver);
            //driver = eventFiringWebDriver;
            // Attaching events             
            // Attaching click events      
            // Element finding events             
            //eventFiringWebDriver.FindingElement += EventFiringWebDriver_FindingElement;
            //eventFiringWebDriver.FindElementCompleted += EventFiringWebDriver_FindElementCompleted;


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
            bool pagina_cargada = false;
            bool existe_elemento = false;
            log_clear();

            try
            {
                //inicializando el nombre del archivo de log generado por el sistema.



                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "*********************************************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Inciando proceso de reporte de pruebas al MINSA");
                DataTable oTableSP = new DataTable();

                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Obtener datos de las pruebas a reportar (sp en labcore).");
                try
                {
                    string sp_datos_minsa = ConfigurationManager.AppSettings["sp_datos_minsa"].ToString();
                    string sp_quantity = ConfigurationManager.AppSettings["sp_quantity"].ToString();

                    SqlConnection cnnSP = new SqlConnection(cnnLABCORE);
                    SqlDataAdapter daSP = new SqlDataAdapter(sp_datos_minsa, cnnSP);
                    daSP.SelectCommand.CommandType = CommandType.StoredProcedure;
                    daSP.SelectCommand.Parameters.Add("@quantity", SqlDbType.Int).Value = int.Parse(sp_quantity);
                    daSP.SelectCommand.CommandTimeout = 60;
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

                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Conectando a MINSA.");


                        pagina_cargada = false;
                        for (int i = 1; i <= 5; i++)
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                            try
                            {
                                if (i == 1)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Conectando a http://190.34.154.91:7050/");
                                    driver.Navigate().GoToUrl("http://190.34.154.91:7050/");
                                }
                            }
                            catch (Exception)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Página http://190.34.154.91:7050/ no responde.");
                                //throw;
                            }

                            System.Threading.Thread.Sleep(30000 + i * i * 1000);
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando si la página cargó correctamente.");
                            if (driver.FindElements(By.Id("username")).Count() > 0)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Página cargó correctamente.");
                                pagina_cargada = true;
                                break;
                            }
                        }
                        if (!pagina_cargada)
                        {
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Después de varios intentos no se pudo cargar la página : http://190.34.154.91:7050/");
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "El programa se cerrará.");

                            driver.Quit();
                            System.Environment.Exit(0);
                            return;
                        }

                        System.Threading.Thread.Sleep(1000);
                        string minsa_user = ConfigurationManager.AppSettings["minsa_user"].ToString();
                        string minsa_pass = ConfigurationManager.AppSettings["minsa_pass"].ToString();
                        driver.FindElement(By.Id("username")).SendKeys(minsa_user);
                        System.Threading.Thread.Sleep(1000);
                        driver.FindElement(By.Id("password")).SendKeys(minsa_pass);
                        driver.FindElement(By.Id("username")).Click();

                        System.Threading.Thread.Sleep(15000);
                        driver.FindElement(By.Id("password")).Click();
                        driver.FindElement(By.Id("password")).SendKeys(Keys.Enter);
                        System.Threading.Thread.Sleep(1000);


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
                                pagina_cargada = false;
                                //driver.Navigate().GoToUrl("http://190.34.154.91:7050/orderentry");
                                for (int i = 1; i <= 10; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    //if (i == 1)
                                    //{
                                    //    System.Threading.Thread.Sleep(10000 + 4 * i * 1000);
                                    //}
                                    if (i == 4)
                                    {
                                        driver.Navigate().GoToUrl("http://190.34.154.91:7050/orderentry");
                                    }

                                    driver.Navigate().GoToUrl("http://190.34.154.91:7050/orderentry");
                                    int i_aux = i <= 10 ? i : 10;
                                    System.Threading.Thread.Sleep(10000 + 4 * i_aux * 1000);
                                    //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

                                    if (driver.FindElements(By.Id("demo_-10_value")).Count() > 0)
                                    {
                                        pagina_cargada = true;
                                        break;
                                    }
                                }
                                if (!pagina_cargada)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Después de varios intentos no se pudo cargar la página de registro: http://190.34.154.91:7050/orderentry");
                                    error_global = true;
                                    break;
                                }
                            }
                            #region datos_iniciales

                            //counter++;
                            string value = "";
                            string state = "";

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application,
                                "l_id: [" + oRow["l_id"].ToString() +
                                "], Cód. estudio: [" + oRow["e_descripcion"].ToString() +
                                "], Orden: [" + oRow["numero_interno"].ToString() +
                                "]");
                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application,
                                "Cédula: [" + oRow["cedula"].ToString() +
                                "], Nombre: [" + oRow["primer_nombre"].ToString() +
                                 "], Apellido: [" + oRow["primer_apellido"].ToString() +
                                "]");


                            bool no_tocar_demograficos = false;
                            bool no_tocar_fecha_nacimiento = false;
                            bool no_tocar_genero = false;
                            bool no_tocar_direcion = false;
                            bool no_tocar_telefono = false;
                            bool no_tocar_correo = false;

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validación de datos. ");
                            string l_id = oRow["l_id"].ToString();
                            string tipo_documento = oRow["tipo_documento"].ToString();
                            string tipo_documento_completo = oRow["tipo_documento_completo"].ToString();
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
                            string corregimiento_numero = oRow["corregimiento_numero"].ToString();
                            string direccion = oRow["direccion"].ToString().Trim();
                            //string persona_contacto = oRow["persona_contacto"].ToString();
                            //string telefono_contacto = oRow["telefono_contacto"].ToString();
                            string correo = oRow["correo"].ToString();
                            string telefono = oRow["telefono"].ToString().Replace("+507", "").Trim();
                            string tipo_de_orden = oRow["tipo_de_orden"].ToString();
                            string tipo_de_orden_completo = oRow["tipo_de_orden_completo"].ToString();
                            string numero_interno = oRow["numero_interno"].ToString();
                            string procedencia_muestra = oRow["procedencia_muestra"].ToString();
                            string procedencia_muestra_completo = oRow["procedencia_muestra_completo"].ToString();
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

                            #region validacion_de_datos_inicial
                            error_on_validation = false;
                            if (String.IsNullOrEmpty(primer_nombre))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "primer_nombre en blanco.");
                                error_on_validation = true;
                            }
                            if (String.IsNullOrEmpty(primer_apellido))
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "primer_apellido en blanco.");
                                error_on_validation = true;
                            }


                            if (error_on_validation)
                            {
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Hay errores en validacion para este estudio, no se puede reportar.  ");

                                //try to update 'laboratorio' to manual report because error_on_validation==true
                                update_labcore_order(l_id, "3");
                                update_labcore_try(l_id); //se añade un intento
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

                                //intertar click en boton deshacer
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscar si está activo el botón 'deshacer'");
                                IWebElement button_deshacer_verificacion1 = null;
                                if (TryFindElement(By.XPath("//*[@ng-click='vm.eventUndo()']"), out button_deshacer_verificacion1))
                                {
                                    if (button_deshacer_verificacion1.Displayed && button_deshacer_verificacion1.Enabled)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshacer activo. ");
                                        try
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en 'Deshacer'.");
                                            button_deshacer_verificacion1.Click();
                                            System.Threading.Thread.Sleep(3000);
                                        }
                                        catch (Exception)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshacer no se pudo cliquear. ");
                                            recargar_pagina = true;
                                            continue;
                                        }
                                    }
                                    //else
                                    //{
                                    //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshacer no esta habilitado. ");
                                    //    recargar_pagina = true;
                                    //    continue;
                                    //}

                                }


                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando boton 'Nuevo'. ");
                                var button_nuevo = driver.FindElement(By.XPath("//*[@ng-click='vm.eventNew()']"));
                                bool next_foreach = false;

                                for (int i = 1; i <= 5; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 1500);
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
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Botón 'Nuevo' no se pudo cliquear. ");
                                            next_foreach = true;
                                        }
                                    }
                                    else { next_foreach = true; }
                                }
                                if (next_foreach)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Despues de 4 intentos no se pudo cliquear botón 'Nuevo'.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Se pasa al siguiente registro. (añandir intento");
                                    update_labcore_try(l_id); //se añade un intento
                                    recargar_pagina = true;
                                    continue;
                                }
                                #endregion

                                #region Validacion_Paciente

                                //System.Threading.Thread.Sleep(1000);
                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_documento.");
                                ////Tipo documento *	demo_-10_value	tipo_documento
                                //driver.FindElement(By.Id("demo_-10_value")).Click();
                                //driver.FindElement(By.Id("demo_-10_value")).SendKeys(tipo_documento + Keys.Enter);
                                int cedula_formato_intento = 1;
                            CedulaFormato:

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscando campo 'tipo de documento'. ");
                                var campo_tipo_documento = driver.FindElement(By.Id("demo_-10_value"));
                                next_foreach = false;

                                for (int i = 1; i <= 5; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    if (i > 1)
                                    {
                                        System.Threading.Thread.Sleep(i * 2000);
                                    }
                                    if ((campo_tipo_documento.Displayed && campo_tipo_documento.Enabled))
                                    {
                                        try
                                        {
                                            campo_tipo_documento.Clear();
                                            campo_tipo_documento.Click();
                                            System.Threading.Thread.Sleep(1000);
                                            //campo_tipo_documento.SendKeys(tipo_documento);
                                            //campo_tipo_documento.SendKeys(Keys.Enter);
                                            campo_tipo_documento.SendKeys(tipo_documento_completo);
                                            campo_tipo_documento.SendKeys(Keys.Enter);
                                            System.Threading.Thread.Sleep(1200);
                                            //driver.FindElement(By.Id("demo_-10_value")).SendKeys(tipo_documento);
                                            next_foreach = false;
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en 'tipo de documento'.");
                                            value = campo_tipo_documento.GetAttribute("value").Trim();
                                            if (value.ToUpper() == "CC. CEDULA" || value.ToUpper() == "PA. PASAPORTE" || value.ToUpper() == "CE. CEDULA EXTRANJERIA")
                                            {
                                                break;

                                            }
                                        }
                                        catch (Exception)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Campo 'tipo de documento' no se pudo cliquear. ");
                                            next_foreach = true;
                                        }
                                    }
                                    else { next_foreach = true; }
                                }
                                if (next_foreach)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Despues de varios intentos no se pudo cliquear botón 'tipo de documento'.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Se pasa al siguiente registro. añadir intento");
                                    update_labcore_try(l_id); //se añade un intento
                                    recargar_pagina = true;
                                    continue;
                                }

                                //System.Threading.Thread.Sleep(900);
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en cedula.");
                                //Cédula *	demo_-100	cedula
                                driver.FindElement(By.Id("demo_-100")).Click();
                                driver.FindElement(By.Id("demo_-100")).Clear();
                                driver.FindElement(By.Id("demo_-100")).SendKeys(cedula);
                                driver.FindElement(By.Id("demo_-100")).SendKeys(Keys.Enter);

                                //bool error_on_cedula = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si los elemntos de la página cargaron (con click en region).");
                                for (int i = 1; i <= 5; i++)
                                {
                                    //REGIÓN *	demo_1_value	region
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 2000);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar formato de cédula es válido");
                                    IWebElement formato_no_valido = null;
                                    bool existe_formato_no_valido = TryFindElement(By.XPath("//span[@class='orderfield_mark_required ng-scope']"), out formato_no_valido);
                                    //existe_formato_no_valido = TryFindElement(By.XPath("//*[contains(., 'Formato no válido')]"), out formato_no_valido);

                                    //driver.FindElement(By.XPath("//*[contains(., 'Formato no válido')]"));
                                    if (existe_formato_no_valido)
                                    {
                                        try
                                        {
                                            string formato_no_valido_value = formato_no_valido.Text;
                                            if (formato_no_valido_value == "Formato no válido")
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Error en formato de cédula.");
                                                if (cedula_formato_intento == 1)
                                                {
                                                    cedula_formato_intento++;
                                                    tipo_documento_completo = "PA. PASAPORTE";
                                                    goto CedulaFormato;
                                                }
                                                else
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Se pasa a registro MANUAL.");
                                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                                    recargar_pagina = true;
                                                    break;
                                                }

                                            }

                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }


                                    //***********************************************************
                                    IWebElement region_verificacion = null;
                                    existe_elemento = TryFindElement(By.Id("demo_1_value"), out region_verificacion); //driver.FindElement(By.Id("demo_-101"));
                                    System.Threading.Thread.Sleep(i * 500 * 2);
                                    if (existe_elemento)
                                    {
                                        next_foreach = false;
                                        if ((region_verificacion.Displayed && region_verificacion.Enabled))
                                        {
                                            try
                                            {

                                                region_verificacion.Click();
                                                System.Threading.Thread.Sleep(300);
                                                recargar_pagina = false;
                                                break;
                                            }
                                            catch (Exception)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Elemento (región) no disponible, seguir intentando.");
                                                recargar_pagina = true;



                                            }

                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Elemento (región) no disponible, seguir intentando.");
                                            recargar_pagina = true;


                                        }
                                        else
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Elemento (región) no habilitado, seguir intentando.");
                                            recargar_pagina = true;
                                        }
                                    }
                                    else
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Elemento (región) no existe, seguir intentando.");
                                        recargar_pagina = true;
                                    }
                                }
                                if (recargar_pagina)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Después de varios intentos no se activaron los elementos.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro. (se intentará de nuevo)");
                                    update_labcore_try(l_id); //se añade un intento
                                    continue;

                                }

                                //validar si boton guardar se activó
                                //validar si el boton de guardar esta disponible

                                //intertar click en boton deshacer
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Buscar si está activo el botón 'guardar'");
                                IWebElement button_guardar_verificacion1 = null;
                                next_foreach = false;
                                if (TryFindElement(By.XPath("//*[@ng-click='vm.eventSave()']"), out button_guardar_verificacion1))
                                {
                                    if (button_guardar_verificacion1.Displayed && button_guardar_verificacion1.Enabled)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Botón guardar activo. ");
                                    }
                                    else
                                    {
                                        next_foreach = true;
                                    }

                                }
                                if (next_foreach)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Botón 'Guardar' NO está inactivo.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Se pasa al siguiente registro. (añandir intento");
                                    update_labcore_try(l_id); //se añade un intento
                                    recargar_pagina = true;
                                    continue;
                                }


                                IWebElement formato_invalido_label = null;
                                bool formato_invalido = TryFindElement(By.XPath("//*[contains(., 'Formato no válido')]"), out formato_invalido_label);
                                if (formato_invalido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Cédula con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual

                                    recargar_pagina = true;

                                    continue;
                                }


                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si existe el paciente.");
                                //validar si existe el paciente
                                //var element = WaitForElement(1, By.Id("demo_-103"), driver);
                                //value = element.GetAttribute("value").Trim();

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar si campo 'Primer Apellido' esta disponible. ");
                                ////Primer apellido *  demo_-101   primer_apellido

                                //bool next_foreach = false;
                                value = "";
                                next_foreach = false;
                                for (int i = 1; i <= 4; i++)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());

                                    IWebElement primer_apellido_verificacion = null;
                                    existe_elemento = TryFindElement(By.Id("demo_-101"), out primer_apellido_verificacion); //driver.FindElement(By.Id("demo_-101"));
                                    System.Threading.Thread.Sleep(i * 500 * 2);
                                    if (existe_elemento)
                                    {
                                        next_foreach = false;
                                        if ((primer_apellido_verificacion.Displayed && !primer_apellido_verificacion.Enabled)) // existe pero esta deshabilitado //paciente existe
                                        {
                                            try
                                            {
                                                //paciente existe
                                                value = primer_apellido_verificacion.GetAttribute("value").Trim();
                                                if (!String.IsNullOrEmpty(value))
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Campo 'Primer Apellido', leído con éxito.");
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Paciente existe.");
                                                    break;
                                                }
                                                else
                                                {
                                                    IWebElement primer_nombre_verificacion = null;
                                                    existe_elemento = TryFindElement(By.Id("demo_-103"), out primer_nombre_verificacion);

                                                    value = primer_nombre_verificacion.GetAttribute("value").Trim();
                                                    if (!String.IsNullOrEmpty(value))
                                                    {
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Campo 'Primer Apellido', leído con éxito.");
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Paciente existe.");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        //valor incoherente: valor vacio con campo inactivo. Puede que la página no ha terminado de cargar
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Campo 'Primer Apellido' / 'Primer nombre': inactivo pero con valor vacío.");
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Pasar al siguiente intento");
                                                        next_foreach = true;
                                                        recargar_pagina = true;
                                                    }

                                                }

                                            }
                                            catch (Exception)
                                            {
                                                next_foreach = true;
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Campo 'Primer apellido' no se pudo leer. ");
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "No se puede continuar con el registro actual.");
                                            }
                                        }
                                        else
                                        {
                                            if ((primer_apellido_verificacion.Displayed && primer_apellido_verificacion.Enabled))//paciente NO existe
                                            {
                                                value = "";
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Paciente no existe.");
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Campo 'Primer Apellido', leído con éxito.");
                                                break;
                                            }
                                            else
                                            {
                                                next_foreach = true;
                                            }
                                        }
                                    }
                                    else { next_foreach = true; }

                                }//if (existe_elemento)

                                ////validar si campo 'primer apellido' esta deshabilitado y en blanco.al parecer nunca ocurre
                                //if (!primer_apellido_verificacion.Enabled and String.IsNullOrEmpty(value)) )
                                //{
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Campo 'Primer apellido' esta deshabilitado pero vacío.");
                                //    continue;
                                //}
                                if (next_foreach)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Despues de 4 intentos no se pudo leer el campo 'Primer apellido'.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "****Se pasa al siguiente registro. (Se intentará en otra corrida)");
                                    update_labcore_try(l_id); //se añade un intento
                                    continue;
                                }

                                //value = driver.FindElement(By.Id("demo_-103")).GetAttribute("value").Trim();

                                if (!String.IsNullOrEmpty(value))
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Paciente existente.");
                                    string comentario_de_la_orden = "";
                                    paciente_existe = true;

                                    ////evauar nombre 1
                                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si primer nombre de labcore es diferente de primer nombre de minsa.");
                                    ////Primer nombre *	demo_-103	primer_nombre
                                    //string[] primer_nombre_values = driver.FindElement(By.Id("demo_-103")).GetAttribute("value").Trim().Split(' ');
                                    //string primer_nombre_minsa = primer_nombre_values[0];

                                    //if (primer_nombre != primer_nombre_minsa)
                                    //{
                                    //    comentario_de_la_orden = "PACIENTE: " + primer_nombre + " " + primer_apellido + "." + Keys.Enter;
                                    //}

                                    ////Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si fecha nacimiento de labcore es diferente de fecha nacimiento de minsa.");
                                    //////evaluar fecha nacimiento
                                    ////string fecha_nacimiento_value = driver.FindElement(By.Id("demo_-105")).GetAttribute("value").Trim();
                                    ////if (fecha_nacimiento != fecha_nacimiento_value)
                                    ////{
                                    ////    comentario_de_la_orden += "FECHA NACIMIENTO: " + fecha_nacimiento + "." + Keys.Enter;
                                    ////}
                                    //if (!String.IsNullOrEmpty(comentario_de_la_orden))
                                    //{
                                    //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en comentario_de_la_orden: " + comentario_de_la_orden);
                                    //    IWebElement commentario_box = null;
                                    //    bool commentario = TryFindElement(By.XPath("//*[@ng-model='commentorder.notes[0].commentArray.content']"), out commentario_box);
                                    //    if (commentario)
                                    //    {
                                    //        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ng-model='commentorder.notes[0].commentArray.content' existe.");
                                    //    }
                                    //    else
                                    //    {
                                    //        commentario = TryFindElement(By.Id("ui-tinymce-6"), out commentario_box);
                                    //        if (commentario)
                                    //        {
                                    //            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ui-tinymce-6' existe.");
                                    //        }
                                    //        else
                                    //        {
                                    //            commentario = TryFindElement(By.Id("ui-tinymce-12"), out commentario_box);
                                    //            if (commentario)
                                    //            {
                                    //                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Comentario textbox 'ui-tinymce-12' existe.");
                                    //            }
                                    //        }
                                    //    }
                                    //    if (commentario)
                                    //    {
                                    //        try
                                    //        {
                                    //            //en caso de no coincidir el nombre añadir nota
                                    //            //Comentario de la orden	ui-tinymce-6	comentario_de_la_orden
                                    //            commentario_box.Click();
                                    //            commentario_box.Clear();
                                    //            commentario_box.SendKeys(comentario_de_la_orden);
                                    //            System.Threading.Thread.Sleep(800);
                                    //        }
                                    //        catch (Exception)
                                    //        {
                                    //            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "****Error al añadir el comentario.");
                                    //            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "***No se pudo añadir el comentario.");
                                    //        }
                                    //    }
                                    //    else
                                    //    {
                                    //        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "***Comentario textbox no hallado.");
                                    //    }
                                    //}
                                }
                                #endregion

                                #region validacion_adicional

                                string genero_completo_aux = GetElementValueById("demo_-104_value").Replace(".", ". ").ToUpper();
                                if (!String.IsNullOrEmpty(genero_completo_aux))
                                {
                                    if (genero_completo_aux == genero_completo)
                                    {
                                        no_tocar_fecha_nacimiento = true;
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(genero_completo))
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "genero en blanco. ");
                                            error_on_validation = true;
                                        }
                                    }

                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(genero_completo))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "genero en blanco. ");
                                        error_on_validation = true;
                                    }

                                }


                                //if (String.IsNullOrEmpty(genero_completo))
                                //{
                                //    genero_completo = GetElementValueById("demo_-104_value");
                                //    if (String.IsNullOrEmpty(genero_completo))
                                //    {
                                //        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "genero en blanco. ");
                                //        error_on_validation = true;
                                //    }
                                //    {
                                //        no_tocar_genero = true;
                                //    }
                                //}

                                //Fecha nacimiento *	demo_-105	fecha_nacimiento
                                string fecha_nacimiento_aux = GetElementValueById("demo_-105");
                                if (!String.IsNullOrEmpty(fecha_nacimiento_aux))
                                {
                                    if (fecha_nacimiento_aux == fecha_nacimiento)
                                    {
                                        no_tocar_fecha_nacimiento = true;
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(fecha_nacimiento))
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha de nacimiento en blanco. ");
                                            error_on_validation = true;
                                        }
                                    }

                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(fecha_nacimiento))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha de nacimiento en blanco. ");
                                        error_on_validation = true;
                                    }

                                }


                                //if (String.IsNullOrEmpty(fecha_nacimiento))
                                //{
                                //    //Fecha nacimiento *	demo_-105	fecha_nacimiento
                                //    fecha_nacimiento = GetElementValueById("demo_-105");
                                //    if (String.IsNullOrEmpty(fecha_nacimiento))
                                //    {
                                //        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_nacimiento en blanco.");
                                //        error_on_validation = true;
                                //    }
                                //    else
                                //    {
                                //        no_tocar_fecha_nacimiento = true;
                                //    }

                                //}


                                //if (String.IsNullOrEmpty(region))
                                //{
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "region en blanco. ");
                                //    error_on_validation = true;
                                //}
                                //if (String.IsNullOrEmpty(distrito))
                                //{
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "distrito en blanco.");
                                //    error_on_validation = true;
                                //    //distrito = "95."; //95. SIN DEFINIR
                                //    //distrito_completo = "95. SIN DEFINIR";
                                //}

                                //CORREGIMIENTO *	demo_3_value	corregimiento
                                string corregimiento_aux = GetElementValueById("demo_3_value").Replace(".", ". ");
                                //DISTRITO *	demo_2_value	distrito
                                string distrito_aux = GetElementValueById("demo_2_value").Replace(".", ". ");
                                //REGIÓN *	demo_1_value	region
                                string region_aux = GetElementValueById("demo_1_value").Replace(".", ". ");
                                if (!String.IsNullOrEmpty(corregimiento_aux) && !String.IsNullOrEmpty(distrito_aux) && !String.IsNullOrEmpty(region_aux))
                                {

                                    //if (corregimiento_completo == corregimiento_aux)
                                    if (corregimiento_aux.Contains(corregimiento_completo))
                                    {
                                        if (region_completo != "14. C. NGOBE BUGLE")
                                        {
                                            no_tocar_demograficos = true;
                                        }
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(corregimiento_completo))
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "corregimiento en blanco.");
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "corregimiento de minsa disponible.");
                                            //error_on_validation = true;
                                        }
                                    }
                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(corregimiento_completo))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "corregimiento en blanco.");
                                        error_on_validation = true;
                                    }
                                }

                                //Dirección	demo_-112	direccion
                                string direccion_aux = GetElementValueById("demo_-112").Trim().ToUpper();
                                direccion_aux = Truncate(direccion_aux, 99);
                                if (!String.IsNullOrEmpty(direccion_aux))
                                {
                                    //if (direccion_aux == direccion)
                                    if (direccion_aux.Contains(direccion))
                                    {
                                        no_tocar_direcion = true;
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(direccion))
                                        {
                                            if (String.IsNullOrEmpty(direccion_aux))
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "direccion en blanco. ");
                                                error_on_validation = true;

                                            }
                                            else
                                            {
                                                no_tocar_direcion = true;
                                            }


                                            //direccion = corregimiento;
                                        }
                                    }

                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(direccion))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "direccion en blanco. ");
                                        error_on_validation = true;
                                        //direccion = corregimiento;
                                    }

                                }

                                //Teléfono	demo_-111	telefono
                                string telefono_aux = GetElementValueById("demo_-111").Trim();
                                if (!String.IsNullOrEmpty(telefono_aux))
                                {
                                    //if (telefono_aux == telefono || String.IsNullOrEmpty(telefono))
                                    if (telefono_aux.Contains(telefono) || String.IsNullOrEmpty(telefono))
                                    {
                                        no_tocar_telefono = true;
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(telefono))
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "telefono en blanco. ");
                                            error_on_validation = true;
                                            //telefono = "No aportó";
                                        }
                                    }

                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(telefono))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "telefono en blanco. ");
                                        error_on_validation = true;
                                        //telefono = "No aportó";
                                    }

                                }

                                //Correo	demo_-106	correo
                                string correo_aux = GetElementValueById("demo_-106").ToUpper();
                                string[] correo_values = correo.Trim().Split(',');
                                string primer_correo = correo_values[0];

                                if (!String.IsNullOrEmpty(correo_aux))
                                {

                                    //if (correo_aux.ToUpper() == primer_correo.ToUpper())
                                    if (correo_aux.Contains(primer_correo.ToUpper()))
                                    {
                                        no_tocar_correo = true;
                                    }
                                    else
                                    {
                                        if (String.IsNullOrEmpty(correo))
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "correo en blanco. ");
                                            //error_on_validation = true;
                                        }
                                    }

                                }
                                else
                                {
                                    if (String.IsNullOrEmpty(correo))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "correo en blanco. ");
                                        //error_on_validation = true;
                                    }

                                }



                                ////if (resultado_valor == "1") //para resultados positivos es obligatorio un telefono
                                ////{
                                //if (String.IsNullOrEmpty(telefono))
                                //{
                                //    //Teléfono	demo_-111	telefono
                                //    telefono = GetElementValueById("demo_-111");
                                //    if (String.IsNullOrEmpty(telefono))
                                //    {
                                //        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "telefono en blanco. ");
                                //        error_on_validation = true;
                                //    }
                                //    else
                                //    {
                                //        no_tocar_telefono = true;
                                //    }
                                //}
                                //}
                                if (error_on_validation)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Hay errores en validacion adicional para este estudio, no se puede reportar.  ");

                                    //try to update 'laboratorio' to manual report because error_on_validation==true
                                    update_labcore_order(l_id, "3");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Pasando a la siguiente orden. ");
                                    recargar_pagina = true;
                                    continue; //next for 
                                }

                                #endregion


                                #region ingreso_de_datos

                                if (!paciente_existe)
                                {


                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en primer_apellido.");
                                    //Primer apellido *	demo_-101	primer_apellido
                                    string primer_apellido_aux = GetElementValueById("demo_-101");
                                    if (primer_apellido_aux != primer_apellido)
                                    {
                                        if (!primer_apellido.Contains(primer_apellido_aux) || String.IsNullOrEmpty(primer_apellido_aux))
                                        //if (!primer_apellido.Contains(primer_apellido_aux ) )
                                        {

                                            var element_primer_apellido = driver.FindElement(By.Id("demo_-101"));
                                            //System.Threading.Thread.Sleep(500);
                                            element_primer_apellido.Click();
                                            element_primer_apellido.Clear();
                                            element_primer_apellido.SendKeys(primer_apellido + Keys.Enter);
                                            System.Threading.Thread.Sleep(300);
                                        }
                                    }

                                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en segundo_apellido.");
                                    ////Segundo apellido	demo_-102	segundo_apellido
                                    //driver.FindElement(By.Id("demo_-102")).Click();
                                    driver.FindElement(By.Id("demo_-102")).Clear();
                                    ////driver.FindElement(By.Id("demo_-102")).SendKeys(segundo_apellido + Keys.Enter);
                                    //System.Threading.Thread.Sleep(300);

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en primer_nombre.");
                                    //Primer nombre *	demo_-103	primer_nombre
                                    string primer_nombre_aux = GetElementValueById("demo_-103");
                                    if (primer_nombre_aux != primer_nombre)
                                    {
                                        if (!primer_nombre.Contains(primer_nombre_aux) || String.IsNullOrEmpty(primer_nombre_aux))
                                        //if (!primer_nombre.Contains(primer_nombre_aux) )
                                        {

                                            driver.FindElement(By.Id("demo_-103")).Click();
                                            driver.FindElement(By.Id("demo_-103")).Clear();
                                            driver.FindElement(By.Id("demo_-103")).SendKeys(primer_nombre + Keys.Enter);
                                            System.Threading.Thread.Sleep(300);
                                        }
                                    }

                                    //Segundo nombre	demo_-109	segundo_nombre
                                    //driver.FindElement(By.Id("demo_-109")).SendKeys(segundo_nombre + Keys.Enter); //no lo usamos

                                    driver.FindElement(By.Id("demo_-109")).Clear();



                                    System.Threading.Thread.Sleep(300);
                                    if (!no_tocar_genero)
                                    {
                                        if (genero_completo != GetElementValueById("demo_-104_value"))
                                        {

                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en genero.");
                                            bool error_on_sex = false;
                                            for (int i = 0; i < 4; i++)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Intentar registrar el género.");
                                                //Género *	demo_-104_value	genero
                                                driver.FindElement(By.Id("demo_-104_value")).Click();
                                                driver.FindElement(By.Id("demo_-104_value")).Clear();
                                                driver.FindElement(By.Id("demo_-104_value")).SendKeys(genero_completo);
                                                System.Threading.Thread.Sleep(i * 200 + 1000);
                                                driver.FindElement(By.Id("demo_-104_value")).SendKeys(Keys.Tab);
                                                System.Threading.Thread.Sleep(300);

                                                string genero_minsa_aux = driver.FindElement(By.Id("demo_-104_value")).GetAttribute("value").Trim();
                                                if (String.IsNullOrEmpty(genero_minsa_aux))
                                                {
                                                    error_on_sex = true;
                                                }
                                                else
                                                {
                                                    error_on_sex = false;
                                                    break;
                                                }
                                            }
                                            if (error_on_sex)
                                            {
                                                recargar_pagina = true;
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error al registrar el género.");
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Pasando al siguiente registro.");
                                                continue;
                                            }
                                        }
                                    }


                                }//if (!paciente_existe)

                                if (!no_tocar_fecha_nacimiento)
                                {
                                    if (fecha_nacimiento != GetElementValueById("demo_-105"))
                                    {

                                        /////////    FECHA_NACIMIENTO
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_nacimiento.");
                                        //Fecha nacimiento *	demo_-105	fecha_nacimiento
                                        driver.FindElement(By.Id("demo_-105")).Click();
                                        driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                        System.Threading.Thread.Sleep(300);
                                        driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                        System.Threading.Thread.Sleep(300);
                                        driver.FindElement(By.Id("demo_-105")).SendKeys(fecha_nacimiento);
                                        //System.Threading.Thread.Sleep(300);


                                        //******************

                                        bool fec_nac_valido = false;
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar fecha de nacimiento.");
                                        for (int i = 1; i <= 5; i++)
                                        {
                                            //REGIÓN *	demo_1_value	region
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                            System.Threading.Thread.Sleep(i * 500);

                                            //***********************************************************
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta la fecha de nacimiento");
                                            IWebElement fecha_nacimiento_val = null;
                                            bool existe_fecha_nac = TryFindElement(By.Id("demo_-105"), out fecha_nacimiento_val);
                                            //existe_formato_no_valido = TryFindElement(By.XPath("//*[contains(., 'Formato no válido')]"), out formato_no_valido);

                                            //driver.FindElement(By.XPath("//*[contains(., 'Formato no válido')]"));
                                            if (existe_fecha_nac)
                                            {
                                                try
                                                {

                                                    DateTime Temp;
                                                    string fecha_nacimiento_val_value = fecha_nacimiento_val.GetAttribute("value");
                                                    if (DateTime.TryParse(fecha_nacimiento_val_value, out Temp) == true)
                                                    {
                                                        fec_nac_valido = true;
                                                    }
                                                    else
                                                    {
                                                        fec_nac_valido = false;
                                                    }
                                                    if (fec_nac_valido)
                                                    {
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fecha de nacimiento correcta.");
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en formato de fecha de nacimiento.");
                                                        driver.FindElement(By.Id("demo_-105")).Click();
                                                        driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                                        System.Threading.Thread.Sleep(300);
                                                        driver.FindElement(By.Id("demo_-105")).SendKeys(Keys.ArrowLeft);
                                                        System.Threading.Thread.Sleep(300);
                                                        driver.FindElement(By.Id("demo_-105")).SendKeys(fecha_nacimiento);
                                                        System.Threading.Thread.Sleep(300);
                                                    }

                                                }
                                                catch (Exception)
                                                {

                                                }

                                            }
                                        }

                                        if (!fec_nac_valido)
                                        {
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fecha de nacimiento con formato inválido, requiere REGISTRO MANUAL.");
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                            update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                            recargar_pagina = true;
                                            continue;
                                        }
                                    }
                                }
                                ///  FIN    FECHA_NACIMIENTO
                                //***********************************************************


                                //******************
                                if (!no_tocar_demograficos)
                                {

                                    //////////////          REGION
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en region.");
                                    region = region_completo;
                                    //REGIÓN *	demo_1_value	region
                                    driver.FindElement(By.Id("demo_1_value")).Clear();
                                    driver.FindElement(By.Id("demo_1_value")).Click();
                                    driver.FindElement(By.Id("demo_1_value")).SendKeys(region);
                                    System.Threading.Thread.Sleep(1200);
                                    driver.FindElement(By.Id("demo_1_value")).SendKeys(Keys.Enter);
                                    System.Threading.Thread.Sleep(200);

                                    //******************

                                    bool region_valido = false;
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar region.");
                                    for (int i = 1; i <= 3; i++)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                        //System.Threading.Thread.Sleep(i * 2000);

                                        //***********************************************************
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta la region");
                                        IWebElement region_val = null;
                                        bool existe_region = TryFindElement(By.Id("demo_1_value"), out region_val);

                                        if (existe_region)
                                        {
                                            try
                                            {
                                                string region_val_value = region_val.GetAttribute("value");
                                                if (!string.IsNullOrEmpty(region_val_value))
                                                {
                                                    region_valido = true;
                                                }
                                                else
                                                {
                                                    region_valido = false;
                                                }
                                                if (region_valido)
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Región correcta.");
                                                    break;
                                                }
                                                else
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de Región.");
                                                    driver.FindElement(By.Id("demo_1_value")).Clear();
                                                    driver.FindElement(By.Id("demo_1_value")).Click();
                                                    driver.FindElement(By.Id("demo_1_value")).SendKeys(region);
                                                    System.Threading.Thread.Sleep(i * 1500);
                                                    driver.FindElement(By.Id("demo_1_value")).SendKeys(Keys.Enter);
                                                    System.Threading.Thread.Sleep(200);
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }
                                        }
                                    }

                                    if (!region_valido)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Región con formato inválido, requiere REGISTRO MANUAL.");
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                        update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                        recargar_pagina = true;
                                        continue;
                                    }
                                    //FIN REGION
                                    //***********************************************************



                                    //////////////          DISTRITO
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en distrito.");
                                    distrito = distrito_completo;
                                    //DISTRITO *	demo_2_value	distrito
                                    driver.FindElement(By.Id("demo_2_value")).Clear();
                                    driver.FindElement(By.Id("demo_2_value")).Click();
                                    driver.FindElement(By.Id("demo_2_value")).SendKeys(distrito);
                                    System.Threading.Thread.Sleep(1200);
                                    driver.FindElement(By.Id("demo_2_value")).SendKeys(Keys.Enter);
                                    System.Threading.Thread.Sleep(200);

                                    //******************

                                    bool distrito_valido = false;
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar distrito.");
                                    for (int i = 1; i <= 3; i++)
                                    {

                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                        //System.Threading.Thread.Sleep(i * 2000);

                                        //***********************************************************
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el distrito");
                                        IWebElement distrito_val = null;
                                        bool existe_distrito = TryFindElement(By.Id("demo_2_value"), out distrito_val);

                                        if (existe_distrito)
                                        {
                                            try
                                            {
                                                string distrito_val_value = distrito_val.GetAttribute("value");
                                                if (!string.IsNullOrEmpty(distrito_val_value))
                                                {
                                                    distrito_valido = true;
                                                }
                                                else
                                                {
                                                    distrito_valido = false;
                                                }
                                                if (distrito_valido)
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Distrito correcto.");
                                                    break;
                                                }
                                                else
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de distrito.");
                                                    driver.FindElement(By.Id("demo_2_value")).Clear();
                                                    driver.FindElement(By.Id("demo_2_value")).Click();
                                                    driver.FindElement(By.Id("demo_2_value")).SendKeys(distrito);
                                                    System.Threading.Thread.Sleep(i * 1500);
                                                    driver.FindElement(By.Id("demo_2_value")).SendKeys(Keys.Enter);
                                                    System.Threading.Thread.Sleep(200);
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }

                                        }
                                    }

                                    if (!distrito_valido)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Distrito con formato inválido, requiere REGISTRO MANUAL.");
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                        update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                        recargar_pagina = true;
                                        continue;
                                    }
                                    //FIN DISTRITO
                                    //***********************************************************



                                    //////////////          CORREGIMIENTO
                                    //string[] words = corregimiento_completo.Split('.');
                                    //corregimiento = words[1].Trim();
                                    corregimiento = corregimiento_completo;
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en corregimiento.");
                                    //CORREGIMIENTO *	demo_3_value	corregimiento
                                    driver.FindElement(By.Id("demo_3_value")).Clear();
                                    driver.FindElement(By.Id("demo_3_value")).Click();
                                    driver.FindElement(By.Id("demo_3_value")).SendKeys(corregimiento);
                                    System.Threading.Thread.Sleep(1300);
                                    driver.FindElement(By.Id("demo_3_value")).SendKeys(Keys.Enter);
                                    System.Threading.Thread.Sleep(200);

                                    //******************

                                    bool corregimiento_valido = false;
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar Corregimiento.");
                                    for (int i = 1; i <= 3; i++)
                                    {

                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                        //System.Threading.Thread.Sleep(i * 2000);

                                        //***********************************************************
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el corregimiento");
                                        IWebElement corregimiento_val = null;
                                        bool existe_corregimiento = TryFindElement(By.Id("demo_3_value"), out corregimiento_val);

                                        if (existe_corregimiento)
                                        {
                                            try
                                            {
                                                string corregimiento_val_value = corregimiento_val.GetAttribute("value");
                                                if (!string.IsNullOrEmpty(corregimiento_val_value))
                                                {
                                                    corregimiento_valido = true;
                                                }
                                                else
                                                {
                                                    corregimiento_valido = false;
                                                }
                                                if (corregimiento_valido)
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Corregimiento correcto.");
                                                    break;
                                                }
                                                else
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de corregimiento.");
                                                    driver.FindElement(By.Id("demo_3_value")).Clear();
                                                    driver.FindElement(By.Id("demo_3_value")).Click();
                                                    driver.FindElement(By.Id("demo_3_value")).SendKeys(corregimiento_numero + ".");
                                                    System.Threading.Thread.Sleep(i * 1500);
                                                    driver.FindElement(By.Id("demo_3_value")).SendKeys(Keys.Enter);
                                                    System.Threading.Thread.Sleep(200);
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }

                                        }
                                    }

                                    if (!corregimiento_valido)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Corregimiento con formato inválido, requiere REGISTRO MANUAL.");
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                        update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                        recargar_pagina = true;
                                        continue;
                                    }
                                    //FIN CORREGIMIENTO
                                    //***********************************************************

                                } //fin no tocar demograficos





                                if (!no_tocar_direcion)
                                {
                                    if (direccion != GetElementValueById("demo_-112"))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en direccion.");
                                        //Dirección	demo_-112	direccion
                                        driver.FindElement(By.Id("demo_-112")).Clear();
                                        driver.FindElement(By.Id("demo_-112")).SendKeys(direccion + Keys.Enter);
                                        System.Threading.Thread.Sleep(200);
                                    }

                                }


                                //PERSONA CONTACTO	demo_5	persona_contacto
                                //TELÉFONO CONTACTO	demo_6	telefono_contacto


                                //correo ELECTRONICO
                                if (!no_tocar_correo)
                                {
                                    if (!String.IsNullOrEmpty(correo))
                                    {
                                        //string[] correo_values = correo.Trim().Split(',');
                                        //string primer_correo = correo_values[0];
                                        //Correo	demo_-106	correo
                                        driver.FindElement(By.Id("demo_-106")).Clear();
                                        driver.FindElement(By.Id("demo_-106")).SendKeys(primer_correo + Keys.Enter);
                                        System.Threading.Thread.Sleep(400);
                                    }
                                }

                                //TELEFONO
                                if (!no_tocar_telefono)
                                {
                                    //if (!String.IsNullOrEmpty(telefono))
                                    //{
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en telefono.");
                                    //Teléfono	demo_-111	telefono
                                    driver.FindElement(By.Id("demo_-111")).Clear();
                                    driver.FindElement(By.Id("demo_-111")).SendKeys(telefono + Keys.Enter);
                                    System.Threading.Thread.Sleep(200);
                                    //}
                                }

                                ////Tipo de orden *	demo_-4_value	tipo_de_orden
                                //driver.FindElement(By.Id("demo_-4_value")).Clear();
                                //driver.FindElement(By.Id("demo_-4_value")).SendKeys(tipo_de_orden_completo + Keys.Tab);
                                ////System.Threading.Thread.Sleep(1000);

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en numero_interno.");
                                //NÚMERO INTERNO *	demo_8	numero_interno
                                driver.FindElement(By.Id("demo_8")).Click();
                                driver.FindElement(By.Id("demo_8")).SendKeys(numero_interno + Keys.Enter);
                                System.Threading.Thread.Sleep(300);

                                //PROCEDENCIA MUESTRA	demo_9_value	procedencia_muestra 
                                driver.FindElement(By.Id("demo_9_value")).Clear();
                                driver.FindElement(By.Id("demo_9_value")).Click();
                                driver.FindElement(By.Id("demo_9_value")).SendKeys(procedencia_muestra_completo);
                                System.Threading.Thread.Sleep(400);
                                driver.FindElement(By.Id("demo_9_value")).SendKeys(Keys.Enter);
                                System.Threading.Thread.Sleep(200);

                                //driver.FindElement(By.Id("demo_9_value")).SendKeys(procedencia_muestra_completo + Keys.Enter);
                                //System.Threading.Thread.Sleep(500);


                                /////////    FECHA_DE_SINTOMAS

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_sintomas.");
                                //FECHA DE SÍNTOMAS *	demo_18	fecha_sintomas
                                driver.FindElement(By.Id("demo_18")).Click();
                                driver.FindElement(By.Id("demo_18")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(300);
                                driver.FindElement(By.Id("demo_18")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(300);
                                driver.FindElement(By.Id("demo_18")).SendKeys(fecha_sintomas);
                                System.Threading.Thread.Sleep(300);

                                //******************

                                bool fecha_de_sintomas_valido = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar fecha_sintomas.");
                                for (int i = 1; i <= 3; i++)
                                {
                                    //REGIÓN *	demo_1_value	region
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 500);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta la fecha_de_sintomas");
                                    IWebElement fecha_de_sintomas_val = null;
                                    bool existe_fecha_de_sintomas = TryFindElement(By.Id("demo_18"), out fecha_de_sintomas_val);
                                    if (existe_fecha_de_sintomas)
                                    {
                                        try
                                        {

                                            DateTime Temp;
                                            string fecha_de_sintomas_val_value = fecha_de_sintomas_val.GetAttribute("value");
                                            if (DateTime.TryParse(fecha_de_sintomas_val_value, out Temp) == true)
                                            {
                                                fecha_de_sintomas_valido = true;
                                            }
                                            else
                                            {
                                                fecha_de_sintomas_valido = false;
                                            }
                                            if (fecha_de_sintomas_valido)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_de_toma correcta.");
                                                break;
                                            }
                                            else
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en formato de fecha_sintomas.");
                                                driver.FindElement(By.Id("demo_18")).Click();
                                                driver.FindElement(By.Id("demo_18")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(200);
                                                driver.FindElement(By.Id("demo_18")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(200);
                                                driver.FindElement(By.Id("demo_18")).SendKeys(fecha_sintomas);
                                                System.Threading.Thread.Sleep(300);
                                            }

                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }
                                }

                                if (!fecha_de_sintomas_valido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_de_sintomas con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                    recargar_pagina = true;
                                    continue;
                                }

                                ///  FIN    FECHA DE SINTOMAS
                                //***********************************************************

                                /////////    FECHA_DE_TOMA
                                ///
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_de_toma.");
                                //FECHA DE TOMA	demo_10	fecha_de_toma 
                                driver.FindElement(By.Id("demo_10")).Click();
                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(200);
                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(200);
                                driver.FindElement(By.Id("demo_10")).SendKeys(fecha_de_toma);
                                System.Threading.Thread.Sleep(300);


                                //******************

                                bool fecha_de_toma_valido = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar fecha_de_toma.");
                                for (int i = 1; i <= 3; i++)
                                {
                                    //REGIÓN *	demo_1_value	region
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 500);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta la fecha_de_toma");
                                    IWebElement fecha_de_toma_val = null;
                                    bool existe_fecha_de_toma = TryFindElement(By.Id("demo_10"), out fecha_de_toma_val);
                                    if (existe_fecha_de_toma)
                                    {
                                        try
                                        {

                                            DateTime Temp;
                                            string fecha_de_toma_val_value = fecha_de_toma_val.GetAttribute("value");
                                            if (DateTime.TryParse(fecha_de_toma_val_value, out Temp) == true)
                                            {
                                                fecha_de_toma_valido = true;
                                            }
                                            else
                                            {
                                                fecha_de_toma_valido = false;
                                            }
                                            if (fecha_de_toma_valido)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_de_toma correcta.");
                                                break;
                                            }
                                            else
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en formato de fecha_de_toma.");
                                                driver.FindElement(By.Id("demo_10")).Click();
                                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(200);
                                                driver.FindElement(By.Id("demo_10")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(200);
                                                driver.FindElement(By.Id("demo_10")).SendKeys(fecha_de_toma);
                                                System.Threading.Thread.Sleep(300);
                                            }

                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }
                                }

                                if (!fecha_de_toma_valido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_de_toma con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                    recargar_pagina = true;
                                    continue;
                                }

                                ///  FIN    FECHA_DE_TOMA
                                //***********************************************************





                                //////////////          TIPO_DE_PRUEBA

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_prueba.");
                                //TIPO DE PRUEBA *	demo_12_value	tipo_de_prueba 
                                driver.FindElement(By.Id("demo_12_value")).Clear();
                                driver.FindElement(By.Id("demo_12_value")).Click();
                                driver.FindElement(By.Id("demo_12_value")).SendKeys(tipo_de_prueba);
                                System.Threading.Thread.Sleep(1000);
                                driver.FindElement(By.Id("demo_12_value")).SendKeys(Keys.Enter);
                                System.Threading.Thread.Sleep(500);

                                //******************


                                bool tipo_de_prueba_valido = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar tipo_de_prueba.");
                                for (int i = 1; i <= 3; i++)
                                {

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    //System.Threading.Thread.Sleep(i * 2000);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el tipo_de_prueba");
                                    IWebElement tipo_de_prueba_val = null;
                                    bool existe_tipo_de_prueba = TryFindElement(By.Id("demo_12_value"), out tipo_de_prueba_val);

                                    if (existe_tipo_de_prueba)
                                    {
                                        try
                                        {
                                            string tipo_de_prueba_val_value = tipo_de_prueba_val.GetAttribute("value");
                                            if (!string.IsNullOrEmpty(tipo_de_prueba_val_value))
                                            {
                                                tipo_de_prueba_valido = true;
                                            }
                                            else
                                            {
                                                tipo_de_prueba_valido = false;
                                            }
                                            if (tipo_de_prueba_valido)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_prueba correcto.");
                                                break;
                                            }
                                            else
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de tipo_de_prueba.");
                                                driver.FindElement(By.Id("demo_12_value")).Clear();
                                                driver.FindElement(By.Id("demo_12_value")).Click();
                                                driver.FindElement(By.Id("demo_12_value")).SendKeys(tipo_de_prueba);
                                                System.Threading.Thread.Sleep(i * 1500);
                                                driver.FindElement(By.Id("demo_12_value")).SendKeys(Keys.Enter);
                                                System.Threading.Thread.Sleep(500);
                                            }
                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }
                                }

                                if (!tipo_de_prueba_valido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_prueba con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                    recargar_pagina = true;
                                    continue;
                                }
                                //FIN TIPO_DE_PRUEBA
                                //***********************************************************



                                //////////////          RESULTADO_MINSA

                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en resultado_minsa.");
                                //RESULTADO *	demo_11_value	resultado_minsa 
                                driver.FindElement(By.Id("demo_11_value")).Clear();
                                driver.FindElement(By.Id("demo_11_value")).Click();
                                driver.FindElement(By.Id("demo_11_value")).SendKeys(resultado_minsa);
                                System.Threading.Thread.Sleep(1000);
                                driver.FindElement(By.Id("demo_11_value")).SendKeys(Keys.Enter);
                                System.Threading.Thread.Sleep(500);

                                //******************

                                bool resultado_minsa_valido = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar resultado_minsa.");
                                for (int i = 1; i <= 3; i++)
                                {

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    //System.Threading.Thread.Sleep(i * 2000);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el resultado_minsa");
                                    IWebElement resultado_minsa_val = null;
                                    bool existe_resultado_minsa = TryFindElement(By.Id("demo_11_value"), out resultado_minsa_val);

                                    if (existe_resultado_minsa)
                                    {
                                        try
                                        {
                                            string resultado_minsa_val_value = resultado_minsa_val.GetAttribute("value");
                                            if (!string.IsNullOrEmpty(resultado_minsa_val_value))
                                            {
                                                resultado_minsa_valido = true;
                                            }
                                            else
                                            {
                                                resultado_minsa_valido = false;
                                            }
                                            if (resultado_minsa_valido)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "resultado_minsa correcto.");
                                                break;
                                            }
                                            else
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de resultado_minsa.");
                                                driver.FindElement(By.Id("demo_11_value")).Clear();
                                                driver.FindElement(By.Id("demo_11_value")).Click();
                                                driver.FindElement(By.Id("demo_11_value")).SendKeys(resultado_minsa);
                                                System.Threading.Thread.Sleep(i * 1500);
                                                driver.FindElement(By.Id("demo_11_value")).SendKeys(Keys.Enter);
                                                System.Threading.Thread.Sleep(500);
                                            }
                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }
                                }

                                if (!resultado_minsa_valido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "resultado_minsa con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                    recargar_pagina = true;
                                    continue;
                                }
                                //FIN RESULTADO_MINSA
                                //***********************************************************




                                //RESULTADO IGG *	demo_16_value	resultado_igg
                                //RESULTADO IGM *	demo_17_value	resultado_igm

                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_resultado.");
                                ////FECHA RESULTADO *	demo_13	fecha_resultado
                                //driver.FindElement(By.Id("demo_13")).Click();
                                //driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                //System.Threading.Thread.Sleep(200);
                                //driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                //System.Threading.Thread.Sleep(900);
                                //driver.FindElement(By.Id("demo_13")).SendKeys(fecha_resultado);
                                //System.Threading.Thread.Sleep(300);

                                /////////    FECHA_RESULTADO
                                ///
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en fecha_resultado.");
                                //FECHA RESULTADO *	demo_13	fecha_resultado
                                driver.FindElement(By.Id("demo_13")).Click();
                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(200);
                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                System.Threading.Thread.Sleep(900);
                                driver.FindElement(By.Id("demo_13")).SendKeys(fecha_resultado);
                                System.Threading.Thread.Sleep(300);


                                //******************

                                bool fecha_resultado_valido = false;
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar fecha_resultado.");
                                for (int i = 1; i <= 3; i++)
                                {
                                    //REGIÓN *	demo_1_value	region
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                    System.Threading.Thread.Sleep(i * 500);

                                    //***********************************************************
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta la fecha_resultado");
                                    IWebElement fecha_resultado_val = null;
                                    bool existe_fecha_resultado = TryFindElement(By.Id("demo_13"), out fecha_resultado_val);
                                    //existe_formato_no_valido = TryFindElement(By.XPath("//*[contains(., 'Formato no válido')]"), out formato_no_valido);

                                    //driver.FindElement(By.XPath("//*[contains(., 'Formato no válido')]"));
                                    if (existe_fecha_resultado)
                                    {
                                        try
                                        {

                                            DateTime Temp;
                                            string fecha_resultado_val_value = fecha_resultado_val.GetAttribute("value");
                                            if (DateTime.TryParse(fecha_resultado_val_value, out Temp) == true)
                                            {
                                                fecha_resultado_valido = true;
                                            }
                                            else
                                            {
                                                fecha_resultado_valido = false;
                                            }
                                            if (fecha_resultado_valido)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_resultado correcta.");
                                                break;
                                            }
                                            else
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en formato de fecha_resultado.");
                                                driver.FindElement(By.Id("demo_13")).Click();
                                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(200);
                                                driver.FindElement(By.Id("demo_13")).SendKeys(Keys.ArrowLeft);
                                                System.Threading.Thread.Sleep(900);
                                                driver.FindElement(By.Id("demo_13")).SendKeys(fecha_resultado);
                                                System.Threading.Thread.Sleep(300);
                                            }

                                        }
                                        catch (Exception)
                                        {

                                        }

                                    }
                                }

                                if (!fecha_resultado_valido)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "fecha_resultado con formato inválido, requiere REGISTRO MANUAL.");
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                    recargar_pagina = true;
                                    continue;
                                }

                                ///  FIN    FECHA_RESULTADO
                                //***********************************************************



                                //try
                                //{
                                //    //intertar colocar tipo_de_paciente
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_paciente.");
                                //    //TIPO DE PACIENTE	demo_14_value	tipo_de_paciente
                                //    driver.FindElement(By.Id("demo_14_value")).Click();
                                //    driver.FindElement(By.Id("demo_14_value")).SendKeys(tipo_de_paciente);
                                //    System.Threading.Thread.Sleep(1000);
                                //    driver.FindElement(By.Id("demo_14_value")).SendKeys(Keys.Enter);
                                //    //System.Threading.Thread.Sleep(400);
                                //}
                                //catch (Exception)
                                //{
                                //    //proceso continua porque no es campo obligatorio
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "****Error al añadir el tipo_de_paciente.");
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "***No se pudo añadir el tipo_de_paciente.");
                                //}
                                //System.Threading.Thread.Sleep(500);



                                ////////////////          TIPO_DE_PACIENTE

                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_paciente.");
                                ////TIPO DE PACIENTE	demo_14_value	tipo_de_paciente
                                //driver.FindElement(By.Id("demo_14_value")).Clear();
                                //driver.FindElement(By.Id("demo_14_value")).Click();
                                //driver.FindElement(By.Id("demo_14_value")).SendKeys(tipo_de_paciente);
                                //System.Threading.Thread.Sleep(1000);
                                //driver.FindElement(By.Id("demo_14_value")).SendKeys(Keys.Enter);
                                //System.Threading.Thread.Sleep(300);

                                ////******************

                                //bool tipo_de_paciente_valido = false;
                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar tipo_de_paciente.");
                                //for (int i = 1; i <= 3; i++)
                                //{

                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                //    //System.Threading.Thread.Sleep(i * 2000);

                                //    //***********************************************************
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el tipo_de_paciente");
                                //    IWebElement tipo_de_paciente_val = null;
                                //    bool existe_tipo_de_paciente = TryFindElement(By.Id("demo_14_value"), out tipo_de_paciente_val);

                                //    if (existe_tipo_de_paciente)
                                //    {
                                //        try
                                //        {
                                //            string tipo_de_paciente_val_value = tipo_de_paciente_val.GetAttribute("value");
                                //            if (!string.IsNullOrEmpty(tipo_de_paciente_val_value))
                                //            {
                                //                tipo_de_paciente_valido = true;
                                //            }
                                //            else
                                //            {
                                //                tipo_de_paciente_valido = false;
                                //            }
                                //            if (tipo_de_paciente_valido)
                                //            {
                                //                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_paciente correcto.");
                                //                break;
                                //            }
                                //            else
                                //            {
                                //                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Error en seleccion de tipo_de_paciente.");
                                //                driver.FindElement(By.Id("demo_14_value")).Clear();
                                //                driver.FindElement(By.Id("demo_14_value")).Click();
                                //                driver.FindElement(By.Id("demo_14_value")).SendKeys(tipo_de_paciente);
                                //                System.Threading.Thread.Sleep(i * 1500);
                                //                driver.FindElement(By.Id("demo_14_value")).SendKeys(Keys.Enter);
                                //                System.Threading.Thread.Sleep(500);
                                //            }
                                //        }
                                //        catch (Exception)
                                //        {

                                //        }

                                //    }
                                //}
                                ////EN CASO DE ERROR SE OMITE EL TIPO DE PACIENTE PORQUE NO ES OBLIGATORIO
                                ////if (!tipo_de_paciente_valido)
                                ////{
                                ////    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_paciente con formato inválido, requiere REGISTRO MANUAL.");
                                ////    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                ////    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                ////    recargar_pagina = true;
                                ////    continue;
                                ////}
                                ////FIN TIPO_DE_PACIENTE
                                ////***********************************************************




                                ////try
                                ////{
                                ////    //intertar colocar tipo_de_muestra
                                ////    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_muestra.");
                                ////    //TIPO DE MUESTRA	demo_15_value	tipo_de_muestra
                                ////    driver.FindElement(By.Id("demo_15_value")).Clear();
                                ////    driver.FindElement(By.Id("demo_15_value")).Click();
                                ////    driver.FindElement(By.Id("demo_15_value")).SendKeys(tipo_de_muestra_completo);
                                ////    System.Threading.Thread.Sleep(1000);
                                ////    driver.FindElement(By.Id("demo_15_value")).SendKeys(Keys.Enter);
                                ////    //System.Threading.Thread.Sleep(500);
                                ////}
                                ////catch (Exception)
                                ////{
                                ////    //proceso continua porque no es campo obligatorio
                                ////    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "****Error al añadir el tipo_de_paciente.");
                                ////    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "***No se pudo añadir el tipo_de_paciente.");
                                ////}
                                ////System.Threading.Thread.Sleep(600);


                                ////////////////          TIPO_DE_MUESTRA

                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_muestra.");
                                ////TIPO DE MUESTRA	demo_15_value	tipo_de_muestra
                                //driver.FindElement(By.Id("demo_15_value")).Clear();
                                //driver.FindElement(By.Id("demo_15_value")).Click();
                                //driver.FindElement(By.Id("demo_15_value")).SendKeys(tipo_de_muestra_completo);
                                //System.Threading.Thread.Sleep(1000);
                                //driver.FindElement(By.Id("demo_15_value")).SendKeys(Keys.Enter);
                                //System.Threading.Thread.Sleep(500);

                                ////******************

                                //bool tipo_de_muestra_valido = false;
                                //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificar tipo_de_muestra.");
                                //for (int i = 1; i <= 3; i++)
                                //{

                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "intento: " + i.ToString());
                                //    //System.Threading.Thread.Sleep(i * 2000);

                                //    //***********************************************************
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si está correcta el tipo_de_muestra");
                                //    IWebElement tipo_de_muestra_val = null;
                                //    bool existe_tipo_de_muestra = TryFindElement(By.Id("demo_15_value"), out tipo_de_muestra_val);

                                //    if (existe_tipo_de_muestra)
                                //    {
                                //        try
                                //        {
                                //            string tipo_de_muestra_val_value = tipo_de_muestra_val.GetAttribute("value");
                                //            if (!string.IsNullOrEmpty(tipo_de_muestra_val_value))
                                //            {
                                //                tipo_de_muestra_valido = true;
                                //            }
                                //            else
                                //            {
                                //                tipo_de_muestra_valido = false;
                                //            }
                                //            if (tipo_de_muestra_valido)
                                //            {
                                //                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_muestra correcto.");
                                //                break;
                                //            }
                                //            else
                                //            {
                                //                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en tipo_de_muestra.");
                                //                //TIPO DE MUESTRA	demo_15_value	tipo_de_muestra
                                //                driver.FindElement(By.Id("demo_15_value")).Clear();
                                //                driver.FindElement(By.Id("demo_15_value")).Click();
                                //                driver.FindElement(By.Id("demo_15_value")).SendKeys(tipo_de_muestra_completo);
                                //                System.Threading.Thread.Sleep(i * 1500);
                                //                driver.FindElement(By.Id("demo_15_value")).SendKeys(Keys.Enter);
                                //                System.Threading.Thread.Sleep(500);
                                //            }
                                //        }
                                //        catch (Exception)
                                //        {

                                //        }

                                //    }
                                //}


                                //EN CASO DE ERROR SE OMITE EL TIPO DE MUESTRA PORQUE NO ES OBLIGATORIO
                                //if (!tipo_de_muestra_valido)
                                //{
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "tipo_de_muestra con formato inválido, requiere REGISTRO MANUAL.");
                                //    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro.");
                                //    update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                //    recargar_pagina = true;
                                //    continue;
                                //}
                                //FIN TIPO_DE_MUESTRA
                                //***********************************************************



                                #endregion

                                #region grabar_datos



                                //grabar registro con ALT + S
                                var button_guardar = driver.FindElement(By.XPath("//*[@ng-click='vm.eventSave()']"));
                                try
                                {
                                    bool error_al_guardar = false;
                                    bool repetido_guardado = false;

                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en guardar. ");
                                    button_guardar.Click();

                                    //buscar si es repetida
                                    System.Threading.Thread.Sleep(1200);
                                    IWebElement div_repetido = null;
                                    if (TryFindElement(By.XPath("//div[contains(text(), 'El número interno ya existe en la base de datos')]"), out div_repetido))
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "orden previamente grbada (repetido)");
                                        error_al_guardar = false;
                                        repetido_guardado = true;
                                        goto RegistroRepetido;
                                    }

                                    if (!repetido_guardado)
                                    {
                                        //validar que registro se guardado
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando si al guardar, el formulario respondió correctamente.");
                                        //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Verificando si boton 'Nuevo' esta desactivado y boton guardar está activo. ");

                                        for (int i = 1; i <= 6; i++)
                                        {

                                            //validar si aparece el boton de seguro si desea guardar?
                                            //***************
                                            //***********************************************************
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Validar si se muestra el modal de seguro qye desea guardar");
                                            IWebElement seguro_desea_guardar = null;
                                            bool existe_seguro_desea_guardar = TryFindElement(By.XPath("//*[@ng-click='vm.confirmationsave()']"), out seguro_desea_guardar);

                                            if (existe_seguro_desea_guardar)
                                            {
                                                try
                                                {

                                                    System.Threading.Thread.Sleep(400);
                                                    seguro_desea_guardar.Click();

                                                    //buscar si es repetida
                                                    System.Threading.Thread.Sleep(1000);
                                                    div_repetido = null;
                                                    if (TryFindElement(By.XPath("//div[contains(text(), 'El número interno ya existe en la base de datos')]"), out div_repetido))
                                                    {
                                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "orden previamente grbada (repetido)");
                                                        error_al_guardar = false;
                                                        repetido_guardado = true;
                                                        goto RegistroRepetido;
                                                        //break;
                                                    }

                                                }
                                                catch (Exception)
                                                {

                                                }

                                            }
                                            //***************

                                            System.Threading.Thread.Sleep(3000 + i * 2000);


                                            //evaluar si boton nuevo esta disabled
                                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Intento: " + i.ToString());
                                            IWebElement button_nuevo_verificacion = null;// driver.FindElement(By.XPath("//*[@ng-click='vm.eventNew()']"));
                                            if (TryFindElement(By.XPath("//*[@ng-click='vm.eventNew()']"), out button_nuevo_verificacion))
                                            {
                                                if ((!button_nuevo_verificacion.Enabled) && button_guardar.Enabled)
                                                {
                                                    error_al_guardar = true;
                                                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "El boton 'Nuevo' esta desactivado (no se ha terminado de guardar).");
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No se ha terminado de guardar.");
                                                    //no se ha terminado de guardar.
                                                    //esperar
                                                }
                                                else
                                                {
                                                    error_al_guardar = false;
                                                    //Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "El boton 'Nuevo' esta activo.");
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Registro guardado, formulario listo para otro registro.");
                                                    break;
                                                }
                                            }

                                            try
                                            {
                                                button_guardar.Click();

                                                //buscar si es repetida
                                                System.Threading.Thread.Sleep(500 * i);
                                                div_repetido = null;
                                                if (TryFindElement(By.XPath("//div[contains(text(), 'El número interno ya existe en la base de datos')]"), out div_repetido))
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "orden previamente grbada (repetido)");
                                                    error_al_guardar = false;
                                                    repetido_guardado = true;
                                                    goto RegistroRepetido;
                                                }
                                            }
                                            catch (Exception)
                                            {

                                            }
                                        }
                                    }
                                RegistroRepetido:

                                    if (error_al_guardar)
                                    {
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "No he habilitó el botón nuevo, se colocará este registro para REVISIÓN MANUAL");
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Puede tratarse de registro duplicado o la página tardó mucho en responder.");
                                        //update_labcore_order(l_id, "3"); //esta orden pasa a reporte manual
                                        update_labcore_try(l_id);

                                        //intertar click en boton deshacer
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Buscar si está activo el botón 'deshacer'");
                                        IWebElement button_deshacer_verificacion = null;
                                        if (TryFindElement(By.XPath("//*[@ng-click='vm.eventUndo()']"), out button_deshacer_verificacion))
                                        {
                                            if (button_deshacer_verificacion.Displayed && button_deshacer_verificacion.Enabled)
                                            {
                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshacer activo. ");
                                                try
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Click en 'Deshacer'.");
                                                    button_deshacer_verificacion.Click();
                                                    System.Threading.Thread.Sleep(5000);
                                                }
                                                catch (Exception)
                                                {
                                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshcaer no se pudo cliquear. ");
                                                    recargar_pagina = true;
                                                }
                                            }
                                            else
                                            {

                                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Botón deshacer no esta habilitado. ");
                                                recargar_pagina = true;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //no hubo error se registra como enviado al minsa.
                                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Datos registrados. ");
                                        if (!repetido_guardado)
                                        {

                                            update_labcore_order(l_id, "1");//esta orden se guarda como nueva
                                        }
                                        else
                                        {
                                            update_labcore_order(l_id, "2"); //esta orden se guarda como repetido
                                        }
                                    }

                                }
                                catch (Exception ex)
                                {
                                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error en click guardar = '" + l_id + "'");
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
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Se pasa al siguiente registro");
                                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                                update_labcore_try(l_id); //se añade un intento
                                recargar_pagina = true;
                                continue; //pasar al suguiente registro
                                //break; //end for
                            }//end try 

                            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin de la orden actual. ");

                        }//end foreach (DataRow oRow in oTableSP.Rows)
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "Fin de las órdenes. ");
                    }
                    else
                    {
                        Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "No hay registros nuevos por insertar");
                    } // end if (oTableSP.Rows.Count > 0)

                } //end if (oTableSP != null)

                //en caso de error enviar correo de advertencia
                //if (error_global)


            }
            catch (Exception ex)
            {
                try
                {
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR ***********************");
                    Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                }
                catch (Exception)
                {

                    //throw;
                }

                error_global = true;
            }

            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "EL PROCESO HA FINALIZADO.");
            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, ".........................");
            if (error_global || error_on_validation || error_on_process)
            {
                Send_Email(error_global, "Alerta: Error en proceso");
            }
            if (error_global)
            {
                Send_Email_Log();
            }

            //driver.Close();

            driver.Quit();
            #endregion


        }
        //end public void start_Process()

        //obtener valor por 

        public string GetElementValueById(string ById)
        {
            string element_value = "";


            Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application, "obtener valor del elemento (By.Id): " + ById);
            IWebElement element = null;
            bool exist_element = TryFindElement(By.Id(ById), out element);

            if (exist_element)
            {
                try
                {
                    element_value = element.GetAttribute("value");
                }
                catch (Exception)
                {
                    element_value = "";
                }

            }
            return element_value;

        }



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
                string file = "email_log_emplate.txt";
                //string file = @"C:\\bot_minsa\\email_log_emplate.txt";
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

                vFile_Attach_Path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Log.txt";
                //vFile_Attach_Path = @"C:\\bot_minsa\\Log.txt";
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
                    file = @"email_error_template.txt";
                    //file = @"C:\\bot_minsa\\email_error_template.txt";
                }
                else
                {
                    file = @"email_log_emplate.txt";
                    //file = @"C:\\bot_minsa\\email_log_emplate.txt";
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

        //public IWebElement WaitForElement(int seconds, By By, IWebDriver driver)
        //{
        //    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));
        //    return wait.Until(drv => drv.FindElement(By));
        //}

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
        public bool update_labcore_try(string l_id)
        {
            try  //update labcore.dbo.laboratorio.l_minsa_enviado = 3 // 3 = validation error
            {
                vSql = "UPDATE "
                        + "LABORATORIOS "
                        + "SET "
                        + "l_minsa_intento = isnull(l_minsa_intento,0) + 1 "
                        + "WHERE "
                        + "l_id = '" + l_id
                        + "'";
                SqlExecuteNonQuery(vSql, cnnLABCORE);
                return true;

            }
            catch (Exception ex)
            {
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Error actualizando el registro: LABCORE.dbo.LABORATORIOS l_id = '" + l_id + "'");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "Consulta: Añadir intento");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
                return false;

            }
        }

        public string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) { return value; }

            return value.Substring(0, Math.Min(value.Length, maxLength));
        }


    }


}


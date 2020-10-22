using Gph_FrameWork.Logger;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace bot_minsa.BPM
{
    class cls_bpm
    {
        private const string userName = "sistemas";
        private const string userPassword = "sistemas2019#";
        private const string serverUri = "https://cordondevida.creatio.com/0/ServiceModel/EntityDataService.svc/";
        private const string authServiceUtri = "https://cordondevida.creatio.com/ServiceModel/AuthService.svc/Login";

        // Links to XML name spaces.
        private static readonly XNamespace ds = "http://schemas.microsoft.com/ado/2007/08/dataservices";
        private static readonly XNamespace dsmd = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
        private static readonly XNamespace atom = "http://www.w3.org/2005/Atom";


        public static string validate_product_data(string ult_act_sap, string codigo, string nombre, string cpt, string precio, string validFor, string frozenFor, string inactivo,
                                                   string categoria, string proveedor, string especialidad, string salud_femenina, string tiempo_de_resultados, string tiempo_entrega_dias,
                                                   string seccion, string tipo_de_prueba, string dias_de_procesamiento, string dias_toma_muestra, string condicion_del_paciente,
                                                   string tubo_para_la_toma, string tipo_de_muestra, string dias_de_envio_tiempo_de_viaje, string info_tecnica_refrigeado,
                                                   string condicion_de_la_muestra, string cantidad_para_analisis, string descripcion, string necesidad_asociada,
                                                   ref string id)
        {
            string result = "";

            // Creating an authentication request.
            var authRequest = HttpWebRequest.Create(authServiceUtri) as HttpWebRequest;
            authRequest.Method = "POST";
            authRequest.ContentType = "application/json";
            authRequest.Timeout = 7000;
            authRequest.ReadWriteTimeout = 20000;
            var bpmCookieContainer = new CookieContainer();
            // Including the cookie use into the request.
            authRequest.CookieContainer = bpmCookieContainer;
            // Receiving a stream associated with the authentication request.
            using (var requestStream = authRequest.GetRequestStream())
            {
                // Recording the BPMonline user accounts and additional request parameters into the stream.
                using (var writer = new StreamWriter(requestStream))
                {
                    writer.Write(@"{
                                ""UserName"":""" + userName + @""",
                                ""UserPassword"":""" + userPassword + @""",
                                ""SolutionName"":""TSBpm"",
                                ""TimeZoneOffset"":-120,
                                ""Language"":""En-us""
                                }");
                }
            }
            // Receiving an answer from the server. If the authentication is successful, cookies will placed in the
            // bpmCookieContainer object and they may be used for further requests.
            using (var response = (HttpWebResponse)authRequest.GetResponse())
            {
                // Creating a request for data reception from the OData service.
                var dataRequest = HttpWebRequest.Create(serverUri + "ProductCollection?$filter=Code eq '" + codigo + "'") as HttpWebRequest;
                dataRequest.Timeout = 7000;
                dataRequest.ReadWriteTimeout = 20000;
                // The HTTP method GET is used to receive data.
                dataRequest.Method = "GET";
                // Adding pre-received authentication cookie to the data receipt request.
                dataRequest.CookieContainer = bpmCookieContainer;
                // Receiving a response from the server.


                using (var dataResponse = (HttpWebResponse)dataRequest.GetResponse())
                {
                    // Uploading the server response to an xml-document for further processing.
                    XDocument xmlDoc = XDocument.Load(dataResponse.GetResponseStream());
                    // Receiving the collection of contact objects that comply with the request condition.
                    var entity_bpm = from entry in xmlDoc.Descendants(atom + "entry")
                                     select new
                                     {
                                         Id = new Guid(entry.Element(atom + "content")
                                             .Element(dsmd + "properties")
                                             .Element(ds + "Id").Value),
                                         Name = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "Name").Value,
                                         Code = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "Code").Value,
                                         Price = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "Price").Value,
                                         CodigoCPT = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "UsrCodigoCPT").Value,
                                         IsArchive = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "IsArchive").Value,
                                         UsrUltimaActualizacionSap = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "UsrUltimaActualizacionSap").Value
                                         //,CategoryId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "CategoryId").Value,
                                         //UsrProveedorId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrProveedorId").Value,
                                         //UsrEspecialidadProductoId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrEspecialidadProductoId").Value,
                                         //UsrSaludFemenina = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrSaludFemenina").Value,
                                         //UsrTiempoResultados = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrTiempoResultados").Value,
                                         ,
                                         UsrTiempoEntregaDias = entry.Element(atom + "content")
                                                         .Element(dsmd + "properties")
                                                         .Element(ds + "UsrTiempoEntregaDias").Value
                                         //UsrSeccionId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrSeccionId").Value,
                                         //UsrTipoPruebaId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrTipoPruebaId").Value,
                                         //UsrDiasProcesamiento = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrDiasProcesamiento").Value,
                                         //UsrDiasTomaMuestra = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrDiasTomaMuestra").Value,
                                         //UsrCondicionPaciente = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrCondicionPaciente").Value,
                                         //UsrTuboToma = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrTuboToma").Value,
                                         //UsrTipoMuestraId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrTipoMuestraId").Value,
                                         //UsrDiasEnvioTViaje = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrDiasEnvioTViaje").Value,
                                         //UsrInfoTecRefrigerado = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrInfoTecRefrigerado").Value,
                                         //UsrCondicionMuestra = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrCondicionMuestra").Value,
                                         //UsrCantidadAnalisis = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrCantidadAnalisis").Value,
                                         //Notes = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "Notes").Value,
                                         //UsrNecesidadId = entry.Element(atom + "content")
                                         //                .Element(dsmd + "properties")
                                         //                .Element(ds + "UsrNecesidadId").Value 
                                     };
                    int q = entity_bpm.Count();
                    Console.WriteLine("Total encontrados " + q);

                    if (q == 0)
                    {
                        // no existe 
                        //se debe insertar
                        result = "insert";

                    }
                    if (q >= 2)
                    {
                        // hay más de un producto se debe corregir 
                        result = "duplicate";
                    }
                    if (q == 1)
                    {
                        id = entity_bpm.FirstOrDefault().Id.ToString();
                        string bpm_last_update_sap = entity_bpm.First().UsrUltimaActualizacionSap.Substring(0, 10);
                        //producto existe. Verificar si hay que actualizarlo (se actualiza si la ultima fecha de actualizacion o el precio son diferentes) 

                        //llevar parse date ult_act_sap == today
                        //llevar parse decimal tiempo_entrega_dias != entity_bpm.First().UsrTiempoEntregaDias.ToString())
                        string TiempotregaDiasBPM = float.Parse(entity_bpm.First().UsrTiempoEntregaDias.ToString()).ToString();
                        tiempo_entrega_dias = float.Parse(tiempo_entrega_dias).ToString();
                        string fecha_sap = "";
                        if (ult_act_sap != "") //la fecha ult act sap tiene valor (se evalua)
                        {
                            fecha_sap = DateTime.ParseExact(ult_act_sap, "yyyy-MM-dd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");
                        }
                        string today = DateTime.Now.ToString("yyyy-MM-dd");



                        if (ult_act_sap != "") //la fecha ult act sap tiene valor (se evalua)
                        {

                            if (ult_act_sap != bpm_last_update_sap
                                || fecha_sap == today
                                || nombre != entity_bpm.First().Name.ToString()
                                || precio != entity_bpm.First().Price.ToString()
                                || (inactivo != entity_bpm.First().IsArchive.ToString())
                                || cpt != entity_bpm.First().CodigoCPT.ToString()
                                || tiempo_entrega_dias != TiempotregaDiasBPM
                               )
                            {
                                //update
                                result = "update";
                            }
                        }
                        //else //la fecha ult act sap esta vacía (no se evalua)
                        //{
                        //    if (nombre != entity_bpm.First().Name.ToString()
                        //        || precio != entity_bpm.First().Price.ToString()
                        //        || (inactivo != entity_bpm.First().IsArchive.ToString())
                        //        || cpt != entity_bpm.First().CodigoCPT.ToString()
                        //        )
                        //    {
                        //        //update
                        //        result = "update";
                        //    }
                        //}

                    }

                    return result;
                }
            }
        }

        public static bool update_product_bpm(string ult_act_sap, string codigo, string nombre, string cpt, string precio, string validFor, string frozenFor, string inactivo,
                                              string categoria, string proveedor, string especialidad, string salud_femenina, string tiempo_de_resultados, string tiempo_entrega_dias,
                                              string seccion, string tipo_de_prueba, string dias_de_procesamiento, string dias_toma_muestra, string condicion_del_paciente,
                                              string tubo_para_la_toma, string tipo_de_muestra, string dias_de_envio_tiempo_de_viaje, string info_tecnica_refrigeado,
                                              string condicion_de_la_muestra, string cantidad_para_analisis, string descripcion, string necesidad_asociada,
                                              string id)
        {
            bool resp = false;

            //a la fecha ult_act_sap se le debe añadir un día porq al actualizar bpm pierde 1 dia (no se sabe porqué)
            if (ult_act_sap != "")
            {
                DateTime date = DateTime.Parse(ult_act_sap);
                ult_act_sap = date.AddDays(1).ToString("yyyy-MM-dd");
            }

            try
            {
                // Id of the object record to be modified.
                string bpm_id = id;
                // Creating an xml message containing data on the modified object.
                var content1 = new XElement(dsmd + "properties",
                        new XElement(ds + "Name", nombre),
                        new XElement(ds + "Code", codigo),
                        new XElement(ds + "Price", precio),
                        new XElement(ds + "PrimaryPrice", precio),
                        new XElement(ds + "IsArchive", inactivo),
                        new XElement(ds + "UsrCodigoCPT", cpt),
                        new XElement(ds + "UsrUltimaActualizacionSap", ult_act_sap),
                        new XElement(ds + "CategoryId", categoria),
                        new XElement(ds + "UsrProveedorId", proveedor),
                        new XElement(ds + "UsrEspecialidadProductoId", especialidad),
                        new XElement(ds + "UsrSaludFemenina", salud_femenina),
                        new XElement(ds + "UsrTiempoResultados", tiempo_de_resultados),
                        new XElement(ds + "UsrTiempoEntregaDias", tiempo_entrega_dias),
                        new XElement(ds + "UsrSeccionId", seccion),
                        new XElement(ds + "UsrTipoPruebaId", tipo_de_prueba),
                        new XElement(ds + "UsrDiasProcesamiento", dias_de_procesamiento),
                        new XElement(ds + "UsrDiasTomaMuestra", dias_toma_muestra),
                        new XElement(ds + "UsrCondicionPaciente", condicion_del_paciente),
                        new XElement(ds + "UsrTuboToma", tubo_para_la_toma),
                        new XElement(ds + "UsrTipoMuestraId", tipo_de_muestra),
                        new XElement(ds + "UsrDiasEnvioTViaje", dias_de_envio_tiempo_de_viaje),
                        new XElement(ds + "UsrInfoTecRefrigerado", info_tecnica_refrigeado),
                        new XElement(ds + "UsrCondicionMuestra", condicion_de_la_muestra),
                        new XElement(ds + "UsrCantidadAnalisis", cantidad_para_analisis),
                        new XElement(ds + "Notes", descripcion),
                        new XElement(ds + "UsrNecesidadId", necesidad_asociada)

                );

                var content2 = new XElement(dsmd + "properties",
                        new XElement(ds + "Name", nombre),
                        new XElement(ds + "Code", codigo),
                        new XElement(ds + "Price", precio),
                        new XElement(ds + "PrimaryPrice", precio),
                        new XElement(ds + "IsArchive", inactivo),
                        new XElement(ds + "UsrCodigoCPT", cpt),
                        new XElement(ds + "CategoryId", categoria),
                        new XElement(ds + "UsrProveedorId", proveedor),
                        new XElement(ds + "UsrEspecialidadProductoId", especialidad),
                        new XElement(ds + "UsrSaludFemenina", salud_femenina),
                        new XElement(ds + "UsrTiempoResultados", tiempo_de_resultados),
                        new XElement(ds + "UsrTiempoEntregaDias", tiempo_entrega_dias),
                        new XElement(ds + "UsrSeccionId", seccion),
                        new XElement(ds + "UsrTipoPruebaId", tipo_de_prueba),
                        new XElement(ds + "UsrDiasProcesamiento", dias_de_procesamiento),
                        new XElement(ds + "UsrDiasTomaMuestra", dias_toma_muestra),
                        new XElement(ds + "UsrCondicionPaciente", condicion_del_paciente),
                        new XElement(ds + "UsrTuboToma", tubo_para_la_toma),
                        new XElement(ds + "UsrTipoMuestraId", tipo_de_muestra),
                        new XElement(ds + "UsrDiasEnvioTViaje", dias_de_envio_tiempo_de_viaje),
                        new XElement(ds + "UsrInfoTecRefrigerado", info_tecnica_refrigeado),
                        new XElement(ds + "UsrCondicionMuestra", condicion_de_la_muestra),
                        new XElement(ds + "UsrCantidadAnalisis", cantidad_para_analisis),
                        new XElement(ds + "Notes", descripcion),
                        new XElement(ds + "UsrNecesidadId", necesidad_asociada)
                );


                //se tienen conten1 y content2 para utilizar el uno o el otro segun 
                //si el producto tiene ultima fecha de actualizacion sap ( si no tiene este campo no se debe añadir)
                var content = (ult_act_sap != "" ? content1 : content2);

                var entry = new XElement(atom + "entry",
                        new XElement(atom + "content",
                                new XAttribute("type", "application/xml"),
                                content)
                        );
                // Creating a request to the service which will modify the object data.
                var request = (HttpWebRequest)HttpWebRequest.Create(serverUri
                        + "ProductCollection(guid'" + bpm_id + "')");
                request.Timeout = 7000;
                request.ReadWriteTimeout = 20000;
                request.Credentials = new NetworkCredential(userName, userPassword);
                // or request.Method = "MERGE";
                request.Method = "PUT";
                request.Accept = "application/atom+xml";
                request.ContentType = "application/atom+xml;type=entry";
                // Recording the xml message to the request stream.
                using (var writer = XmlWriter.Create(request.GetRequestStream()))
                {
                    entry.WriteTo(writer);
                }
                // Receiving a response from the service regarding the operation implementation result.
                using (WebResponse response = request.GetResponse())
                {
                    // Processing the operation implementation result.
                    resp = true;
                }

            }
            catch (Exception ex)
            {
                resp = false;
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR ACTUALIZANDO***********************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }

            return resp;

        }

        public static bool insert_product_bpm(string ult_act_sap, string codigo, string nombre, string cpt, string precio, string validFor, string frozenFor, string inactivo,
                                              string categoria, string proveedor, string especialidad, string salud_femenina, string tiempo_de_resultados, string tiempo_entrega_dias,
                                              string seccion, string tipo_de_prueba, string dias_de_procesamiento, string dias_toma_muestra, string condicion_del_paciente,
                                              string tubo_para_la_toma, string tipo_de_muestra, string dias_de_envio_tiempo_de_viaje, string info_tecnica_refrigeado,
                                              string condicion_de_la_muestra, string cantidad_para_analisis, string descripcion, string necesidad_asociada

            )
        {
            bool resp = false;
            //a la fecha ult_act_sap se le debe añadir un día porq al actualizar bpm pierde 1 dia (no se sabe porqué)
            if (ult_act_sap != "")
            {
                DateTime date = DateTime.Parse(ult_act_sap);
                ult_act_sap = date.AddDays(1).ToString("yyyy-MM-dd");
            }
            try
            {

                // Creating a xml message containing data on the created object.
                var content1 = new XElement(dsmd + "properties",
                          new XElement(ds + "Name", nombre),
                          new XElement(ds + "Code", codigo),
                          new XElement(ds + "Price", precio),
                          new XElement(ds + "PrimaryPrice", precio),
                          new XElement(ds + "UsrCodigoCPT", cpt),
                          new XElement(ds + "UsrUltimaActualizacionSap", ult_act_sap),
                          new XElement(ds + "CategoryId", categoria),
                          new XElement(ds + "UsrProveedorId", proveedor),
                          new XElement(ds + "UsrEspecialidadProductoId", especialidad),
                          new XElement(ds + "UsrSaludFemenina", salud_femenina),
                          new XElement(ds + "UsrTiempoResultados", tiempo_de_resultados),
                          new XElement(ds + "UsrTiempoEntregaDias", tiempo_entrega_dias),
                          new XElement(ds + "UsrSeccionId", seccion),
                          new XElement(ds + "UsrTipoPruebaId", tipo_de_prueba),
                          new XElement(ds + "UsrDiasProcesamiento", dias_de_procesamiento),
                          new XElement(ds + "UsrDiasTomaMuestra", dias_toma_muestra),
                          new XElement(ds + "UsrCondicionPaciente", condicion_del_paciente),
                          new XElement(ds + "UsrTuboToma", tubo_para_la_toma),
                          new XElement(ds + "UsrTipoMuestraId", tipo_de_muestra),
                          new XElement(ds + "UsrDiasEnvioTViaje", dias_de_envio_tiempo_de_viaje),
                          new XElement(ds + "UsrInfoTecRefrigerado", info_tecnica_refrigeado),
                          new XElement(ds + "UsrCondicionMuestra", condicion_de_la_muestra),
                          new XElement(ds + "UsrCantidadAnalisis", cantidad_para_analisis),
                          new XElement(ds + "Notes", descripcion),
                          new XElement(ds + "UsrNecesidadId", necesidad_asociada)

            );
                var content2 = new XElement(dsmd + "properties",
                              new XElement(ds + "Name", nombre),
                              new XElement(ds + "Code", codigo),
                              new XElement(ds + "Price", precio),
                              new XElement(ds + "PrimaryPrice", precio),
                              new XElement(ds + "UsrCodigoCPT", cpt),
                              new XElement(ds + "CategoryId", categoria),
                              new XElement(ds + "UsrProveedorId", proveedor),
                              new XElement(ds + "UsrEspecialidadProductoId", especialidad),
                              new XElement(ds + "UsrSaludFemenina", salud_femenina),
                              new XElement(ds + "UsrTiempoResultados", tiempo_de_resultados),
                              new XElement(ds + "UsrTiempoEntregaDias", tiempo_entrega_dias),
                              new XElement(ds + "UsrSeccionId", seccion),
                              new XElement(ds + "UsrTipoPruebaId", tipo_de_prueba),
                              new XElement(ds + "UsrDiasProcesamiento", dias_de_procesamiento),
                              new XElement(ds + "UsrDiasTomaMuestra", dias_toma_muestra),
                              new XElement(ds + "UsrCondicionPaciente", condicion_del_paciente),
                              new XElement(ds + "UsrTuboToma", tubo_para_la_toma),
                              new XElement(ds + "UsrTipoMuestraId", tipo_de_muestra),
                              new XElement(ds + "UsrDiasEnvioTViaje", dias_de_envio_tiempo_de_viaje),
                              new XElement(ds + "UsrInfoTecRefrigerado", info_tecnica_refrigeado),
                              new XElement(ds + "UsrCondicionMuestra", condicion_de_la_muestra),
                              new XElement(ds + "UsrCantidadAnalisis", cantidad_para_analisis),
                              new XElement(ds + "Notes", descripcion),
                              new XElement(ds + "UsrNecesidadId", necesidad_asociada)

                );

                //se tienen conten1 y content2 para utilizar el uno o el otro segun 
                //si el producto tiene ultima fecha de actualizacion sap ( si no tiene este campo no se debe añadir)
                var content = (ult_act_sap != "" ? content1 : content2);

                var entry = new XElement(atom + "entry",
                            new XElement(atom + "content",
                            new XAttribute("type", "application/xml"), content));
                //Console.WriteLine(entry.ToString());
                // Creating a request to the service which will add a new object to the contacts collection.
                var request = (HttpWebRequest)HttpWebRequest.Create(serverUri + "ProductCollection/");
                request.Timeout = 5000;
                request.ReadWriteTimeout = 20000;
                request.Credentials = new NetworkCredential(userName, userPassword);
                request.Method = "POST";
                request.Accept = "application/atom+xml";
                request.ContentType = "application/atom+xml;type=entry";
                // Recording the xml message to the request stream.
                using (var writer = XmlWriter.Create(request.GetRequestStream()))
                {
                    entry.WriteTo(writer);
                }
                // Receiving a response from the service regarding the operation implementation result.
                using (WebResponse response = request.GetResponse())
                {
                    if (((HttpWebResponse)response).StatusCode == HttpStatusCode.Created)
                    {
                        // Processing the operation implementation result.
                        resp = true;
                    }
                }
            }
            catch (Exception ex)
            {
                resp = false;
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR INSERTANDO ***********************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }

            return resp;
        }


    }





}

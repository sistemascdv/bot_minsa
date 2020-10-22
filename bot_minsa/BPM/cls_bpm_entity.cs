using Gph_FrameWork.Logger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace bot_minsa.BPM
{
    class cls_bpm_entity
    {
        private const string userName = "sistemas";
        private const string userPassword = "sistemas2019#";
        private const string serverUri = "https://cordondevida.creatio.com/0/ServiceModel/EntityDataService.svc/";
        private const string authServiceUtri = "https://cordondevida.creatio.com/ServiceModel/AuthService.svc/Login";

        // Links to XML name spaces.
        private static readonly XNamespace ds = "http://schemas.microsoft.com/ado/2007/08/dataservices";
        private static readonly XNamespace dsmd = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";
        private static readonly XNamespace atom = "http://www.w3.org/2005/Atom";



        public static string validate_data(string nombre, ref string id)
        {
            string result = "";

            // Creating an authentication request.
            var authRequest = HttpWebRequest.Create(authServiceUtri) as HttpWebRequest;
            authRequest.Method = "POST";
            authRequest.ContentType = "application/json";
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
                var dataRequest = HttpWebRequest.Create(serverUri + "AccountCollection?$filter=Name eq '" + nombre + "'") as HttpWebRequest;

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
                                                         .Element(ds + "Name").Value
                                     };
                    int q = entity_bpm.Count();
                    // Console.WriteLine("Total encontrados " + q);

                    if (q == 0)
                    {
                        // no existe 
                        //se debe insertar
                        Console.WriteLine("Proveedor no existe se procede a crearlo.");

                        //insertar 
                        result = "insert";

                    }
                    //if (q >= 2)
                    //{
                    //    // hay más de un producto se debe corregir 
                    //    result = "duplicate";
                    //}
                    if (q == 1)
                    {
                        Console.WriteLine("El proveedor existe.");
                        id = entity_bpm.FirstOrDefault().Id.ToString();

                        result = "";
                    }

                    return result;
                }
            }
        }



        public static string select_bpm_id(string entity_name, string filter_name, string filter_value)
        {
            string result = "";
            try
            {

                // Creating an authentication request.
                var authRequest = HttpWebRequest.Create(authServiceUtri) as HttpWebRequest;
                authRequest.Method = "POST";
                authRequest.ContentType = "application/json";
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
                    var dataRequest = HttpWebRequest.Create(serverUri + entity_name + "Collection?$filter=" + filter_name + " eq '" + filter_value + "'") as HttpWebRequest;

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
                                                             .Element(ds + "Name").Value
                                         };
                        int q = entity_bpm.Count();
                        // Console.WriteLine("Total encontrados " + q);

                        if (q == 0)
                        {
                            //Console.WriteLine("No existe: " + entity_name + "."+ filter_name+" = " + filter_value +".");
                            result = "";
                        }
                        else
                        {
                            //Console.WriteLine("Encontrado: " + entity_name + "." + filter_name + " = " + filter_value + ".");
                            result = entity_bpm.FirstOrDefault().Id.ToString();
                        }


                    }
                }



            }
            catch (Exception ex)
            {
                result = "";
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, "ERROR INSERTANDO ***********************");
                Cls_Logger.WriteToLog_and_Console(Cls_Logger.MessageType.Application_Error, ex.Message.ToString());
            }
            return result;
        }

 

        public static string replaceSpecialCharacters(string attribute)
        {
            //se utiliza esta funcion para reemplazar caracteres especiales.
            //Si hay uno de estos caracteres en la peticion odata (en el url) 
            //no funcionará bien la consulta de buscar correctamente

            //fuente:
            //https://stackoverflow.com/questions/4229054/how-are-special-characters-handled-in-an-odata-query
            //https://web.archive.org/web/20150101222238/http://msdn.microsoft.com/en-us/library/aa226544(SQL.80).aspx
            // replace the single quotes
            attribute = attribute.Replace("'", "''");
            attribute = attribute.Replace("%", "%25");
            attribute = attribute.Replace("+", "%2B");
            attribute = attribute.Replace(@"\", "%2F");
            attribute = attribute.Replace("?", "%3F");
            attribute = attribute.Replace("#", "%23");
            attribute = attribute.Replace("&", "%26");
            return attribute;
        }


    }


}







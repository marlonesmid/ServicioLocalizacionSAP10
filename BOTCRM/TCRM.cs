using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using RestSharp;
using System.Net;
using System.Net.Security;
using System.IO;
using Funciones;
using System.Xml;
using SAPbobsCOM;

namespace BOTCRM
{
    public class TCRM
    {

        public void ActualizaTRMSAP()
        {
            string GetDateNow = DateTime.Now.ToString("yyyy-M-dd");

            string ResponseXMLTRM =  ConsultaTRMSuperfinanciera(DateTime.Now.ToString());

        }

        private string ConsultaTRMSuperfinanciera(string sGetDateNow)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                string XMLRequest = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:act=\"http://action.trm.services.generic.action.superfinanciera.nexura.sc.com.co/\"> <soapenv:Header/>   <soapenv:Body>      <act:queryTCRM>      <tcrmQueryAssociatedDate>%GetDateNow%</tcrmQueryAssociatedDate>      </act:queryTCRM>   </soapenv:Body></soapenv:Envelope>";
                XMLRequest = XMLRequest.Replace("%GetDateNow%", sGetDateNow);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://www.superfinanciera.gov.co/SuperfinancieraWebServiceTRM/TCRMServicesWebService/TCRMServicesWebService?wsdl");
                byte[] bytes;
                bytes = System.Text.Encoding.ASCII.GetBytes(XMLRequest);
                request.ContentType = "text/xml; encoding='utf-8'";
                request.ContentLength = bytes.Length;
                request.Method = "POST";

                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;

                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                using (WebClient webClient = new WebClient())
                {
                    var stream = webClient.OpenRead("https://www.superfinanciera.gov.co/SuperfinancieraWebServiceTRM/TCRMServicesWebService/TCRMServicesWebService?wsdl");
                    using (StreamReader sr = new StreamReader(stream))
                    {
                        var page = sr.ReadToEnd();
                    }
                }

                Stream requestStream = request.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
                HttpWebResponse response;
                response = (HttpWebResponse)request.GetResponse();

                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    string responseStr = new StreamReader(responseStream).ReadToEnd();
                    return responseStr;
                }
                else
                {
                    return null;
                }

            }
            catch (Exception ex)
            {
                #region Consulta Ruta del Log

                XmlDocument xmlQuerys = new XmlDocument();
                xmlQuerys.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

                string PathFileLog = xmlQuerys.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

                #endregion

                DllFunciones.Logger(ex.ToString(), PathFileLog);
                return null;
                
            }

        }

        private void ActualizaTRMSAPBusinessOne(SAPbobsCOM.Company _oCompany)
        {
            SAPbobsCOM.SBObob oExchagueRate =(SAPbobsCOM.SBObob)_oCompany.GetBusinessObject(BoObjectTypes.BoBridge);

            //oExchagueRate.SetCurrencyRate("USD",)
                
        }

    }
}

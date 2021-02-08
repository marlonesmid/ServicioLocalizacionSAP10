using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using BOConnection;
using BOeBillingService;
using BOTCRM;
using Funciones;
using SAPbobsCOM;

namespace ServicioLocalizacion
{
    class CallDlls
    {
        #region Instanciacion DLL
        
        BOConnection.Connection DllConnection = new BOConnection.Connection();
        BOeBillingService.eBillingService DlleBillingService = new BOeBillingService.eBillingService();
        Funciones.Comunes DllFunciones = new Funciones.Comunes();
        BOTCRM.TCRM DllTcrm = new BOTCRM.TCRM();
        
        #endregion

        public void DllsMetodos(object sender, EventArgs e)
        {
            try
            {

                #region Creacion de Objetos

                SAPbobsCOM.Company oCompany;

                #endregion

                #region Consulta Ruta del Log

                XmlDocument xmlQuerys = new XmlDocument();
                xmlQuerys.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

                #region Obtener nombre del archivo

                DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
                string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
                string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
                string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

                #endregion

                string PathFileLog = xmlQuerys.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

                PathFileLog = PathFileLog + "\\"+ sNombreArchivoLog   ;

                #endregion

                #region Consulta cantidad de base de datos y las coloca en un arreglo

                #region Consulta cantidad de base de datos

                XmlNodeList Xnodos = xmlQuerys.GetElementsByTagName("Conexion");
                int CountDataBases = Xnodos.Count;
                int Contador = 0;

                #endregion

                #region Coloca las bases de datos en un arreglo

                string[] DataBases = new string[CountDataBases];

                using (XmlReader reader = XmlReader.Create(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml"))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement())
                        {
                            //return only when you have START tag  
                            switch (reader.Name.ToString())
                            {
                                case "CompanyDB":
                                    DataBases[Contador] = reader.ReadString();
                                    Contador++;

                                    break;
                            }
                        }
                    }
                }


                #endregion

                #endregion

                #region Valida que funcionalidades estan activas

                string sTCRM = xmlQuerys.SelectSingleNode("Configuration/Funcionalidades/TRCM").InnerText;
                sTCRM = sTCRM.Trim();

                string seBillingService = xmlQuerys.SelectSingleNode("Configuration/Funcionalidades/eBillingService").InnerText;
                seBillingService = seBillingService.Trim();

                #endregion

                #region Funciondalidad TRM - tasa representativa del mercado

                if (sTCRM == "SI")
                {
                    #region Valida si es hora de ejecutar el servicio

                    bool bEjecutarTRM = true; 
                     
                    string sHoraEjecucion = xmlQuerys.SelectSingleNode("Configuration/Funcionalidades/TRCM/HoraActualizacion").InnerText;
                    var vHoraEjecucion = sHoraEjecucion.Trim();

                    DateTime DateInitial = DateTime.Parse(vHoraEjecucion, System.Globalization.CultureInfo.InvariantCulture);
                    DateInitial = DateInitial.AddDays(1);

                    DateTime DateExecFinal = DateInitial.AddHours(2);
                    DateExecFinal = DateExecFinal.AddDays(1);

                    DateTime DateExecActual = DateTime.Now;
                    DateExecActual = DateExecActual.AddDays(1);

                    if (DateExecActual.ToUniversalTime() >= DateInitial.ToUniversalTime() && DateExecActual.ToUniversalTime() <= DateExecFinal.ToUniversalTime())
                    {
                        bEjecutarTRM = true;
                    }
                    else
                    {
                        bEjecutarTRM = false;
                    }

                    #endregion

                    if (bEjecutarTRM == true)
                    {
                        for (int i = 0; i < CountDataBases; i++)
                        {
                            #region Establece conexion a SAP Business One

                            oCompany = (SAPbobsCOM.Company)DllConnection.SetApplication(DataBases[i]);

                            #endregion

                            DllTcrm.ActualizaTRMSAP(oCompany, PathFileLog);

                            oCompany.Disconnect();

                            DllFunciones.Logger("Desconectado correctamente de la base de datos: " + DataBases[i], PathFileLog);

                        }
                    }
                }

                #endregion

                #region Servicio eBilling de facturacion electronica automatica

                if (seBillingService == "SI")
                {
                    for (int i = 0; i < CountDataBases; i++)
                    {
                        #region Establece conexion a SAP Business One

                        oCompany = (SAPbobsCOM.Company)DllConnection.SetApplication(DataBases[i]);

                        #endregion

                        #region Ejecuta servicio 

                        DllFunciones.Logger("Servicio de facturacion electronica activo ", PathFileLog);

                        if (oCompany.Connected == true)
                        {
                            DllFunciones.Logger("Buscando documentos para enviar a la DIAN ", PathFileLog);

                            DlleBillingService.EnviarDocumentosDIANServicioLocalizacion(oCompany, PathFileLog);

                            oCompany.Disconnect();

                            DllFunciones.Logger("Desconectado correctamente de la base de datos: " + DataBases[i], PathFileLog);

                        }
                        else
                        {

                            DllFunciones.Logger(" No esta conectado a SAP Business One, por lo cual no se ejecutara el servicio eBilling ", PathFileLog);

                        }

                        #endregion
                    }
                }
                    
                #endregion

            }
            catch (Exception ex)
            {

                #region Consulta Ruta del Log

                XmlDocument xmlQuerys = new XmlDocument();
                xmlQuerys.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

                #region Obtener nombre del archivo

                DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
                string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
                string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
                string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

                #endregion

                string PathFileLog = xmlQuerys.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

                PathFileLog = PathFileLog + "\\" + sNombreArchivoLog;

                #endregion

                DllFunciones.Logger(ex.ToString(), PathFileLog);
            }


        }


    }
}

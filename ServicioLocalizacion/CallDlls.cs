using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using BOConnection;
using BOeBillingService;
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

                #region Valida que funcionalidades estan activas

                string sTCRM = xmlQuerys.SelectSingleNode("Configuration/Funcionalidades/TRCM").InnerText;
                sTCRM = sTCRM.Trim();

                string seBillingService = xmlQuerys.SelectSingleNode("Configuration/Funcionalidades/eBillingService").InnerText;
                seBillingService = seBillingService.Trim();
                
                #endregion

                #region Funciondalidad TRM - tasa representativa del mercado

                if (  sTCRM == "SI")
                {

                }

                #endregion

                #region Servicio eBilling de facturacion electronica automatica

                if (seBillingService == "SI")
                {
                    #region Establece conexion a SAP Business One

                    oCompany = (SAPbobsCOM.Company)DllConnection.SetApplication();

                    #endregion

                    DllFunciones.Logger("Servicio de facturacion electronica activo ", PathFileLog);

                    if (oCompany.Connected == true)
                    {
                        DllFunciones.Logger("Buscando documentos para enviar a la DIAN ", PathFileLog);

                        DlleBillingService.ActualizarEstadoDocumentos(oCompany, PathFileLog);

                        oCompany.Disconnect();

                        DllFunciones.Logger("Desconectado de SAP Business One ", PathFileLog);

                    }
                    else
                    {

                        DllFunciones.Logger(" No esta conectado a SAP Business One, por lo cual no se ejecutara el servicio eBilling ", PathFileLog);

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

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Xml;
using System.IO;

namespace ServicioLocalizacion
{
    public partial class Service1 : ServiceBase
    {
        #region Variables 

        string sPathLog = string.Empty;

        Timer Tiempo = null;

        CallDlls CallDlls = new CallDlls();
        Funciones.Comunes DllFunciones = new Funciones.Comunes();

        #endregion

        public Service1()
        {
            
            #region Consulta de configuración 

            XmlDocument xmlTimer = new XmlDocument();
            xmlTimer.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

            int Intervalo = Convert.ToInt32(xmlTimer.SelectSingleNode("Configuration/Timer/Interval").InnerText);
            sPathLog = xmlTimer.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

            #region Obtener nombre del archivo

            DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
            string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
            string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
            string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

            #endregion

            sPathLog = sPathLog + "\\" + sNombreArchivoLog;


            #endregion

            #region Consulta de tiempo de ejecucion del servicio

            Intervalo = Intervalo * 60000;

            #endregion

            InitializeComponent();
            Tiempo = new Timer();
            Tiempo.Interval = Intervalo;
            Tiempo.Enabled = true;
            Tiempo.Elapsed += new ElapsedEventHandler(CallDlls.DllsMetodos);
        }

        protected override void OnStart(string[] args)
        {
            #region Consulta Path Log

            XmlDocument xmlTimer = new XmlDocument();
            xmlTimer.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

            sPathLog = xmlTimer.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

            #region Obtener nombre del archivo

            DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
            string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
            string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
            string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

            #endregion

            sPathLog = sPathLog + "\\" + sNombreArchivoLog;


            #endregion

            DllFunciones.Logger("Se inicia el servicio Localización Basis One", sPathLog);
        }

        protected override void OnStop()
        {
            #region Consulta Path Log 

            XmlDocument xmlTimer = new XmlDocument();
            xmlTimer.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

            sPathLog = xmlTimer.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

            #region Obtener nombre del archivo

            DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
            string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
            string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
            string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

            #endregion

            sPathLog = sPathLog + "\\" + sNombreArchivoLog;

            #endregion

            DllFunciones.Logger("Se detiene el servicio Localización Basis One", sPathLog);
        }
    }
}

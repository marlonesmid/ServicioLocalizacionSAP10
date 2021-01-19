using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Xml;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Funciones;

namespace BOConnection
{
    public class Connection
    {

        Funciones.Comunes DllFunciones = new Funciones.Comunes();

        public SAPbobsCOM.Company SetApplication()
        {
            #region Varibles y Objetos

            int RsConnect = 0;
            string PathFileLog = null;

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            XmlDocument xmlConfig = new XmlDocument();
            xmlConfig.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

            #endregion

            #region Parametros de conexion

            oCompany.CompanyDB = xmlConfig.SelectSingleNode("Configuration/Conexion/CompanyDB").InnerText;
            oCompany.Server = xmlConfig.SelectSingleNode("Configuration/Conexion/Server").InnerText;
            oCompany.LicenseServer = xmlConfig.SelectSingleNode("Configuration/Conexion/LicenseServer").InnerText;
            oCompany.SLDServer = xmlConfig.SelectSingleNode("Configuration/Conexion/SLDServer").InnerText;
            oCompany.DbUserName = xmlConfig.SelectSingleNode("Configuration/Conexion/DbUserName").InnerText;
            oCompany.DbPassword = xmlConfig.SelectSingleNode("Configuration/Conexion/DbPassword").InnerText;
            oCompany.UserName = xmlConfig.SelectSingleNode("Configuration/Conexion/UserName").InnerText;
            oCompany.Password = xmlConfig.SelectSingleNode("Configuration/Conexion/Password").InnerText;

            if (xmlConfig.SelectSingleNode("Configuration/Conexion/DbServerType").InnerText == "dst_MSSQL2014")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            }
            else if (xmlConfig.SelectSingleNode("Configuration/Conexion/DbServerType").InnerText == "dst_MSSQL2016")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            }
            else if (xmlConfig.SelectSingleNode("Configuration/Conexion/DbServerType").InnerText == "dst_MSSQL2017")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
            }
            else if (xmlConfig.SelectSingleNode("Configuration/Conexion/DbServerType").InnerText == "dst_HANADB")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            }

            #region Obtener nombre del archivo

            DateTime NombreArchivo = Convert.ToDateTime(DateTime.Now);
            string sFecha = NombreArchivo.ToString("yyyy_MM_dd");//fecha
            string sHora = NombreArchivo.ToString("HH_mm_ss");//hora
            string sNombreArchivoLog = "Log " + sFecha + " " + sHora + ".txt";

            #endregion

            PathFileLog = xmlConfig.SelectSingleNode("Configuration/PathLog/PathFile").InnerText;

            PathFileLog = PathFileLog + "\\" + sNombreArchivoLog;

            #endregion

            oCompany.UseTrusted = false;

            RsConnect = oCompany.Connect();

            if (RsConnect != 0)
            {
                DllFunciones.Logger(Convert.ToString(oCompany.GetLastErrorCode()) + oCompany.GetLastErrorDescription(), PathFileLog);
            }
            else
            {
                DllFunciones.Logger("Conectado a SAP de " + oCompany.CompanyName + " Correctamente", PathFileLog);
            }
            return (SAPbobsCOM.Company)oCompany;
        }

        public bool UpdateB2C(string SqlQuery)
        {

            string connetionString = GetConnectionString();
            SqlConnection connection;
            SqlCommand command;

            connection = new SqlConnection(connetionString);
            try
            {
                connection.Open();
                command = new SqlCommand(SqlQuery, connection);
                SqlDataAdapter sqlDataAdap = new SqlDataAdapter(command);

                command.ExecuteNonQuery();

                command.Dispose();
                connection.Close();

                return true;

            }
            catch (Exception)
            {
                return false;
            }

        }

        public DataTable GetDataTable(String SqlQuery)
        {
            DataTable SqlDataTable = new DataTable();

            using (SqlConnection SQlConn = new SqlConnection())
            {
                try
                {
                    SQlConn.ConnectionString = GetConnectionString();
                    SQlConn.Open();
                    SqlDataAdapter SQLDa = new SqlDataAdapter(SqlQuery, SQlConn);
                    SQLDa.Fill(SqlDataTable);
                    SQlConn.Close();
                    return SqlDataTable;
                }
                catch
                {
                    SQlConn.Close();
                    return null;
                }
            }
        }

        public static string GetConnectionString()
        {
            //Creacion Variables 
            string str;

            XmlDocument _xmlConfig = new XmlDocument();
            _xmlConfig.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\Config.xml");

            str = "Data Source=" + _xmlConfig.SelectSingleNode("Configuration/Conexion/Server").InnerText + ";Initial Catalog=" + _xmlConfig.SelectSingleNode("Configuration/ConexionB2C/CompanyDB").InnerText + ";User ID=" + _xmlConfig.SelectSingleNode("Configuration/Conexion/DbUserName").InnerText + ";Password=" + _xmlConfig.SelectSingleNode("Configuration/Conexion/DbPassword").InnerText + "";
            return str;

        }
    }
}

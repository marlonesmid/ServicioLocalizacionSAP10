using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using Funciones;
using System.Reflection;


namespace BOJournalEntrys
{
    public class JournalEntrys
    {
        public void ActualizarTercero(SAPbobsCOM.Company _oCompany, string sPath)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Consulta version Dll

                Assembly Assembly = Assembly.LoadFrom("BOJournalEntrys.dll");
                Version vVersion = Assembly.GetName().Version;

                string sVersionDll = "SERLOC " + vVersion.ToString();

                #endregion

                #region Creacion variables y objetos

                SAPbobsCOM.Recordset oGetJounalEntrysHead = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetJounalEntrysLines = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Consulta asientos a actualizar

                string sGetJounalEntrysHead = DllFunciones.GetStringXMLDocument(_oCompany, "BOJournalEntrys", "Queries", "GetJounalEntrysHead");
                string sGetJounalEntrysLines = DllFunciones.GetStringXMLDocument(_oCompany, "BOJournalEntrys", "Queries", "GetJounalEntrysLines");

                oGetJounalEntrysHead.DoQuery(sGetJounalEntrysHead);

                #endregion

                #region Actualiza asiento por asiento

                if (oGetJounalEntrysHead.RecordCount > 0)
                {
                    oGetJounalEntrysHead.MoveFirst();

                    do
                    {

                        string sGetJounalEntrysLinesCopia = sGetJounalEntrysLines.Replace("%TransId%", Convert.ToString(oGetJounalEntrysHead.Fields.Item(0).Value));

                        oGetJounalEntrysLines.DoQuery(sGetJounalEntrysLinesCopia);

                        SAPbobsCOM.JournalEntries oJournalEntrie = (SAPbobsCOM.JournalEntries)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        int iTransId = Convert.ToInt32(oGetJounalEntrysHead.Fields.Item(0).Value);

                        oJournalEntrie.GetByKey(iTransId);

                        do
                        {
                            oJournalEntrie.Lines.SetCurrentLine(Convert.ToInt32(oGetJounalEntrysLines.Fields.Item(1).Value));
                            oJournalEntrie.Lines.UserFields.Fields.Item("U_BO_SocioNegocio").Value = Convert.ToString(oGetJounalEntrysLines.Fields.Item(3).Value);
                            oJournalEntrie.Lines.UserFields.Fields.Item("U_BO_FlagTer").Value = "Y";

                            oGetJounalEntrysLines.MoveNext();

                        } while (oGetJounalEntrysLines.EoF == false);

                        oJournalEntrie.UserFields.Fields.Item("U_BO_Version").Value = sVersionDll;

                        int Rsd = oJournalEntrie.Update();

                        if (Rsd == 0)
                        {
                            DllFunciones.Logger("Asiento Contable No. " + Convert.ToString(iTransId) + " Actualizado correctamente ", sPath);
                        }
                        else
                        {
                            DllFunciones.Logger(_oCompany.GetLastErrorDescription(), sPath);
                        }

                        oGetJounalEntrysHead.MoveNext();

                    } while (oGetJounalEntrysHead.EoF == false);
                }

                #endregion
                
            }
            catch (Exception e)
            {
                DllFunciones.Logger(e.ToString(), sPath);                
            }
        }
    }
}


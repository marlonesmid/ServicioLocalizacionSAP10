using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using Funciones;
using BOeBillingService.ServicioEmisionFE;
using BOeBillingService.ServicioAdjuntosFE;
using System.IO;
using System.Xml.Serialization;
using System.ServiceModel;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using CrystalDecisions.Windows.Forms;


namespace BOeBillingService
{
    public class eBillingService
    {

        #region Variables Globales

        int Rsd = 0;

        #endregion

        #region Parametros globales TFHKA

        ServicioEmisionFE.ServiceClient serviceClient;
        ServicioAdjuntosFE.ServiceClient serviceClientAdjuntos;

        //Especifica el puerto
        BasicHttpBinding port = new BasicHttpBinding();

        #endregion

        public void eBillingBO()
        {

        }

        private FacturaGeneral oBuillInvoice(SAPbobsCOM.Recordset oCabecera, SAPbobsCOM.Recordset oLineas, SAPbobsCOM.Recordset oImpuestos, SAPbobsCOM.Recordset oImpuestosTotales, SAPbobsCOM.Recordset OCUFEInvoice, string ___TipoDocumento)
        {

            #region Datos Generales Factura

            FacturaGeneral FacturadeVenta = new FacturaGeneral();

            FacturadeVenta.cantidadDecimales = Convert.ToString(oCabecera.Fields.Item("cantidadDecimales").Value.ToString());
            FacturadeVenta.moneda = Convert.ToString(oCabecera.Fields.Item("Moneda").Value.ToString());
            FacturadeVenta.rangoNumeracion = Convert.ToString(oCabecera.Fields.Item("RangoNmeracion").Value.ToString());

            string sredondeoAplicado = Convert.ToString(oCabecera.Fields.Item("RedondeoAplicado").Value.ToString());
            FacturadeVenta.redondeoAplicado = sredondeoAplicado.Replace(",", ".");

            FacturadeVenta.tipoDocumento = Convert.ToString(oCabecera.Fields.Item("tipoDocumento").Value.ToString());
            FacturadeVenta.tipoOperacion = Convert.ToString(oCabecera.Fields.Item("tipoOperacion").Value.ToString());

            string stotalBaseImponible = Convert.ToString(oCabecera.Fields.Item("totalBaseImponible").Value.ToString());
            FacturadeVenta.totalBaseImponible = stotalBaseImponible.Replace(",", ".");

            string stotalBrutoConImpuesto = Convert.ToString(oCabecera.Fields.Item("totalBrutoConImpuesto").Value.ToString());
            FacturadeVenta.totalBrutoConImpuesto = stotalBrutoConImpuesto.Replace(",", ".");

            string stotalMonto = Convert.ToString(oCabecera.Fields.Item("totalMonto").Value.ToString());
            FacturadeVenta.totalMonto = stotalMonto.Replace(",", ".");

            string stotalProductos = Convert.ToString(oCabecera.Fields.Item("totalProductos").Value.ToString());
            FacturadeVenta.totalProductos = stotalProductos.Replace(",", ".");

            string stotalSinImpuestos = Convert.ToString(oCabecera.Fields.Item("totalSinImpuestos").Value.ToString());
            FacturadeVenta.totalSinImpuestos = stotalSinImpuestos.Replace(",", ".");

            #region Descuento generales de la factura

            string ValorDescuentoGeneral = Convert.ToString(oCabecera.Fields.Item("Desc_monto").Value.ToString());
            ValorDescuentoGeneral = ValorDescuentoGeneral.Replace(",", ".");

            if (ValorDescuentoGeneral == "0" || ValorDescuentoGeneral == "0.0" || ValorDescuentoGeneral == "0.00" || ValorDescuentoGeneral == "0.000" || ValorDescuentoGeneral == "0.0000" || ValorDescuentoGeneral == "0.00000" || ValorDescuentoGeneral == "0.000000")
            {

            }
            else
            {
                FacturadeVenta.cargosDescuentos = new CargosDescuentos[1];

                CargosDescuentos DescuentoGeneral = new CargosDescuentos();

                DescuentoGeneral.codigo = Convert.ToString(oCabecera.Fields.Item("Desc_Codigo").Value.ToString());
                DescuentoGeneral.descripcion = Convert.ToString(oCabecera.Fields.Item("Desc_descripcion").Value.ToString());
                DescuentoGeneral.indicador = Convert.ToString(oCabecera.Fields.Item("Desc_indicador").Value.ToString());

                string sDesc_monto = Convert.ToString(oCabecera.Fields.Item("Desc_monto").Value.ToString());
                DescuentoGeneral.monto = sDesc_monto.Replace(",", ".");

                string sDesc_montoBase = Convert.ToString(oCabecera.Fields.Item("Desc_montoBase").Value.ToString());
                DescuentoGeneral.montoBase = sDesc_montoBase.Replace(",", ".");

                string sDesc_porcentaje = Convert.ToString(oCabecera.Fields.Item("Desc_porcentaje").Value.ToString());
                DescuentoGeneral.porcentaje = sDesc_porcentaje.Replace(",", ".");

                DescuentoGeneral.secuencia = Convert.ToString(oCabecera.Fields.Item("Desc_secuencia").Value.ToString());

                FacturadeVenta.cargosDescuentos[0] = DescuentoGeneral;

                FacturadeVenta.totalDescuentos = sDesc_monto.Replace(",", ".");
            }

            #endregion

            FacturadeVenta.consecutivoDocumento = Convert.ToString(oCabecera.Fields.Item("consecutivoDocumento").Value.ToString());
            FacturadeVenta.fechaEmision = oCabecera.Fields.Item("fechaEmision").Value.ToString();

            #endregion

            #region cliente

            Cliente cliente = new Cliente();

            cliente.actividadEconomicaCIIU = Convert.ToString(oCabecera.Fields.Item("actividadEconomicaCIIU").Value.ToString());

            cliente.destinatario = new Destinatario[1];
            Destinatario destinatario1 = new Destinatario();

            destinatario1.canalDeEntrega = Convert.ToString(oCabecera.Fields.Item("canalDeEntrega").Value.ToString());

            #region Revision Correos a Enviar

            #region Variables Correo

            string CorreoDeEntrega1 = Convert.ToString(oCabecera.Fields.Item("correoEntrega1").Value.ToString());
            string CorreoDeEntrega2 = Convert.ToString(oCabecera.Fields.Item("correoEntrega2").Value.ToString());
            string CorreoDeEntrega3 = Convert.ToString(oCabecera.Fields.Item("correoEntrega3").Value.ToString());
            string CorreoDeEntrega4 = Convert.ToString(oCabecera.Fields.Item("correoEntrega4").Value.ToString());
            string CorreoDeEntrega5 = Convert.ToString(oCabecera.Fields.Item("correoEntrega5").Value.ToString());

            int ContadorCorreos = 0;

            #endregion

            #region Contador de los correos a enviar 

            if (string.IsNullOrEmpty(CorreoDeEntrega1))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega2))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega3))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega4))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            if (string.IsNullOrEmpty(CorreoDeEntrega5))
            {

            }
            else
            {
                ContadorCorreos++;
            }

            #endregion

            string[] correoEntrega = new string[ContadorCorreos];

            #region Asignacion de los correos a enviar

            if (ContadorCorreos == 1)
            {
                correoEntrega[0] = CorreoDeEntrega1;
            }
            else if (ContadorCorreos == 2)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
            }
            else if (ContadorCorreos == 3)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
            }
            else if (ContadorCorreos == 4)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
                correoEntrega[3] = CorreoDeEntrega4;
            }
            else if (ContadorCorreos == 5)
            {
                correoEntrega[0] = CorreoDeEntrega1;
                correoEntrega[1] = CorreoDeEntrega2;
                correoEntrega[2] = CorreoDeEntrega3;
                correoEntrega[3] = CorreoDeEntrega4;
                correoEntrega[4] = CorreoDeEntrega5;
            }

            #endregion

            #endregion

            destinatario1.email = correoEntrega;
            destinatario1.fechaProgramada = Convert.ToString(oCabecera.Fields.Item("fechaProgramada").Value.ToString());
            destinatario1.nitProveedorReceptor = Convert.ToString(oCabecera.Fields.Item("nitProveedorReceptor").Value.ToString());
            destinatario1.telefono = Convert.ToString(oCabecera.Fields.Item("telefono").Value.ToString());
            cliente.destinatario[0] = destinatario1;

            cliente.detallesTributarios = new Tributos[1];
            Tributos tributos1 = new Tributos();
            tributos1.codigoImpuesto = Convert.ToString(oCabecera.Fields.Item("codigoImpuesto").Value.ToString());
            cliente.detallesTributarios[0] = tributos1;

            Direccion direccionFiscal = new Direccion();
            direccionFiscal.ciudad = Convert.ToString(oCabecera.Fields.Item("ciudad").Value.ToString());
            direccionFiscal.codigoDepartamento = Convert.ToString(oCabecera.Fields.Item("codigoDepartamento").Value.ToString());
            direccionFiscal.departamento = Convert.ToString(oCabecera.Fields.Item("departamento").Value.ToString());
            direccionFiscal.direccion = Convert.ToString(oCabecera.Fields.Item("direccion").Value.ToString());
            direccionFiscal.lenguaje = Convert.ToString(oCabecera.Fields.Item("lenguaje").Value.ToString());
            direccionFiscal.municipio = Convert.ToString(oCabecera.Fields.Item("municipio").Value.ToString());
            direccionFiscal.pais = Convert.ToString(oCabecera.Fields.Item("pais").Value.ToString());
            direccionFiscal.zonaPostal = Convert.ToString(oCabecera.Fields.Item("zonaPostal").Value.ToString());
            cliente.direccionFiscal = direccionFiscal;
            cliente.direccionCliente = direccionFiscal;

            cliente.email = Convert.ToString(oCabecera.Fields.Item("email").Value.ToString());

            InformacionLegal informacionLegalCliente = new InformacionLegal();
            informacionLegalCliente.codigoEstablecimiento = Convert.ToString(oCabecera.Fields.Item("codigoEstablecimiento").Value.ToString());
            informacionLegalCliente.nombreRegistroRUT = Convert.ToString(oCabecera.Fields.Item("nombreRegistroRUT").Value.ToString());
            informacionLegalCliente.numeroIdentificacion = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacion").Value.ToString());
            informacionLegalCliente.numeroIdentificacionDV = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacionDV").Value.ToString());
            informacionLegalCliente.tipoIdentificacion = Convert.ToString(oCabecera.Fields.Item("tipoIdentificacion").Value.ToString());
            cliente.informacionLegalCliente = informacionLegalCliente;

            cliente.nombreRazonSocial = Convert.ToString(oCabecera.Fields.Item("nombreRazonSocial").Value.ToString());
            cliente.nombreComercial = Convert.ToString(oCabecera.Fields.Item("nombreRazonSocial").Value.ToString());
            cliente.notificar = Convert.ToString(oCabecera.Fields.Item("notificar").Value.ToString());
            cliente.numeroDocumento = Convert.ToString(oCabecera.Fields.Item("numeroDocumento").Value.ToString());
            cliente.numeroIdentificacionDV = Convert.ToString(oCabecera.Fields.Item("numeroIdentificacionDV").Value.ToString());

            cliente.responsabilidadesRut = new Obligaciones[1];
            Obligaciones obligaciones1 = new Obligaciones();
            obligaciones1.obligaciones = Convert.ToString(oCabecera.Fields.Item("obligaciones").Value.ToString());
            obligaciones1.regimen = Convert.ToString(oCabecera.Fields.Item("regimen").Value.ToString());
            cliente.responsabilidadesRut[0] = obligaciones1;

            cliente.tipoIdentificacion = Convert.ToString(oCabecera.Fields.Item("tipoIdentificacion").Value.ToString());
            cliente.tipoPersona = Convert.ToString(oCabecera.Fields.Item("tipoPersona").Value.ToString());

            FacturadeVenta.cliente = cliente;

            #endregion 

            #region Consulta las lineas en la factura de venta

            int CantidadArticulos;
            int SecuenciaArreglo;
            int Posicion;
            CantidadArticulos = oLineas.RecordCount;

            #endregion

            #region Si existen Lineas Asigna los valores de cada columna al arreglo Detalle Factura

            if (CantidadArticulos > 0)
            {
                FacturadeVenta.detalleDeFactura = new FacturaDetalle[CantidadArticulos];

                #region Asignacion Articulos

                oLineas.MoveFirst();

                SecuenciaArreglo = 0;
                Posicion = SecuenciaArreglo + 1;

                do
                {

                    FacturaDetalle Articulo = new FacturaDetalle();

                    #region Detalle articulo

                    Articulo.cantidadPorEmpaque = Convert.ToString(oLineas.Fields.Item("cantidadPorEmpaque").Value.ToString());

                    string scantidadReal = Convert.ToString(oLineas.Fields.Item("cantidadReal").Value.ToString());
                    Articulo.cantidadReal = scantidadReal.Replace(",", ".");

                    Articulo.cantidadRealUnidadMedida = Convert.ToString(oLineas.Fields.Item("cantidadRealUnidadMedida").Value.ToString());

                    string scantidadUnidades = Convert.ToString(oLineas.Fields.Item("cantidadUnidades").Value.ToString());
                    Articulo.cantidadUnidades = scantidadUnidades.Replace(",", ".");

                    Articulo.codigoIdentificadorPais = null;
                    Articulo.codigoProducto = Convert.ToString(oLineas.Fields.Item("codigoProducto").Value.ToString());
                    Articulo.descripcion = Convert.ToString(oLineas.Fields.Item("descripcion").Value.ToString());
                    Articulo.descripcionTecnica = Convert.ToString(oLineas.Fields.Item("descripcion").Value.ToString());
                    Articulo.estandarCodigo = Convert.ToString(oLineas.Fields.Item("estandarCodigo").Value.ToString());
                    Articulo.estandarCodigoProducto = Convert.ToString(oLineas.Fields.Item("estandarCodigoProducto").Value.ToString());

                    #endregion

                    #region Descuentos a nivel de linea

                    string ValorDescuentoLinea = Convert.ToString(oLineas.Fields.Item("Desc_porcentaje").Value);
                    ValorDescuentoLinea = ValorDescuentoLinea.Replace(",", ".");

                    if (ValorDescuentoLinea == "0" || ValorDescuentoLinea == "0.0" || ValorDescuentoLinea == "0.00" || ValorDescuentoLinea == "0.000" || ValorDescuentoLinea == "0.0000" || ValorDescuentoLinea == "0.00000" || ValorDescuentoLinea == "0.000000")
                    {

                    }
                    else
                    {
                        Articulo.cargosDescuentos = new CargosDescuentos[1];

                        CargosDescuentos DescuentoLinea = new CargosDescuentos();

                        DescuentoLinea.descripcion = Convert.ToString(oLineas.Fields.Item("Desc_descripcion").Value.ToString());
                        DescuentoLinea.indicador = Convert.ToString(oLineas.Fields.Item("Desc_indicador").Value.ToString());

                        string sDesc_monto = Convert.ToString(oLineas.Fields.Item("Desc_monto").Value.ToString());
                        DescuentoLinea.monto = sDesc_monto.Replace(",", ".");

                        string sDesc_montoBase = Convert.ToString(oLineas.Fields.Item("Desc_montoBase").Value.ToString());
                        DescuentoLinea.montoBase = sDesc_montoBase.Replace(",", ".");

                        string sPorcentaje = Convert.ToString(oLineas.Fields.Item("Desc_porcentaje").Value.ToString());
                        DescuentoLinea.porcentaje = sPorcentaje.Replace(",", ".");

                        DescuentoLinea.secuencia = Convert.ToString(oLineas.Fields.Item("Desc_secuencia").Value.ToString());

                        Articulo.cargosDescuentos[0] = DescuentoLinea;

                    }

                    #endregion

                    Articulo.impuestosDetalles = new FacturaImpuestos[1];

                    FacturaImpuestos Impuesto = new FacturaImpuestos();

                    #region Detalle Impuesto

                    string sbaseImponibleTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("baseImponibleTOTALImp").Value.ToString());
                    Impuesto.baseImponibleTOTALImp = sbaseImponibleTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.codigoTOTALImp = Convert.ToString(oLineas.Fields.Item("codigoTOTALImp").Value.ToString());
                    Impuesto.controlInterno = Convert.ToString(oLineas.Fields.Item("controlInterno").Value.ToString());

                    string sporcentajeTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("porcentajeTOTALImp").Value.ToString());
                    Impuesto.porcentajeTOTALImp = sporcentajeTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.unidadMedida = Convert.ToString(oLineas.Fields.Item("unidadMedida").Value.ToString());
                    Impuesto.unidadMedidaTributo = Convert.ToString(oLineas.Fields.Item("unidadMedidaTributo").Value.ToString());

                    string svalorTOTALImp_Impuesto = Convert.ToString(oLineas.Fields.Item("valorTOTALImp").Value.ToString());
                    Impuesto.valorTOTALImp = svalorTOTALImp_Impuesto.Replace(",", ".");

                    Impuesto.valorTributoUnidad = Convert.ToString(oLineas.Fields.Item("valorTributoUnidad").Value.ToString());

                    #endregion

                    Articulo.impuestosDetalles[0] = Impuesto;

                    Articulo.impuestosTotales = new ImpuestosTotales[1];

                    ImpuestosTotales ImpuestoTOTAL = new ImpuestosTotales();

                    #region Detalle Impuesto Total

                    ImpuestoTOTAL.codigoTOTALImp = Convert.ToString(oLineas.Fields.Item("codigoTOTALImp").Value.ToString());

                    string smontoTotal_ImpuestoTOTAL = Convert.ToString(oLineas.Fields.Item("montoTotal").Value.ToString());
                    ImpuestoTOTAL.montoTotal = smontoTotal_ImpuestoTOTAL.Replace(",", ".");


                    #endregion

                    Articulo.impuestosTotales[0] = ImpuestoTOTAL;

                    #region Demas detalles de la linea del articulo


                    Articulo.marca = Convert.ToString(oLineas.Fields.Item("marca").Value.ToString());
                    Articulo.muestraGratis = Convert.ToString(oLineas.Fields.Item("muestraGratis").Value.ToString());

                    #region Si la linea del articulo es muestra, coloca el tag precio de referencia


                    if (Articulo.muestraGratis == "1")
                    {
                        Articulo.precioReferencia = Convert.ToString(oLineas.Fields.Item("precioReferencia").Value.ToString());
                    }

                    #endregion

                    string sprecioTotal_Articulo = Convert.ToString(oLineas.Fields.Item("precioTotal").Value.ToString());
                    Articulo.precioTotal = sprecioTotal_Articulo.Replace(",", ".");

                    string sprecioTotalSinImpuestos_Articulo = Convert.ToString(oLineas.Fields.Item("precioTotalSinImpuestos").Value.ToString());
                    Articulo.precioTotalSinImpuestos = sprecioTotalSinImpuestos_Articulo.Replace(",", ".");

                    string sprecioVentaUnitario_Articulo = Convert.ToString(oLineas.Fields.Item("precioVentaUnitario").Value.ToString());
                    Articulo.precioVentaUnitario = sprecioVentaUnitario_Articulo.Replace(",", ".");

                    Articulo.secuencia = Convert.ToString(oLineas.Fields.Item("Secuencia").Value.ToString());
                    Articulo.unidadMedida = Convert.ToString(oLineas.Fields.Item("unidadMedida").Value.ToString());


                    #endregion

                    FacturadeVenta.detalleDeFactura[SecuenciaArreglo] = Articulo;

                    SecuenciaArreglo = SecuenciaArreglo + 1;

                    Posicion = Posicion + 1;

                    oLineas.MoveNext();

                } while (oLineas.EoF == false);

                #endregion
            }

            #endregion

            #region Documento Referenciado

            if (___TipoDocumento == "NotaCreditoClientes" || ___TipoDocumento == "NotaDebitoClientes")
            {
                if (Convert.ToString(oCabecera.Fields.Item("tipoOperacion").Value.ToString()) == "22")
                {

                }
                else
                {

                    #region Arreglo donde se asigna comentarios acerca del motivo de la devolucion o anulacion

                    string[] descripcion = new string[1];

                    descripcion[0] = Convert.ToString(oCabecera.Fields.Item("Comentarios_NC").Value.ToString()); ;

                    #endregion

                    FacturadeVenta.documentosReferenciados = new DocumentoReferenciado[2];

                    #region Documento Referenciado 1

                    DocumentoReferenciado Datos_NCoND_DR1 = new DocumentoReferenciado();

                    Datos_NCoND_DR1.codigoEstatusDocumento = Convert.ToString(oCabecera.Fields.Item("codigoEstatusDocumento").Value.ToString());
                    Datos_NCoND_DR1.codigoInterno = "4";
                    Datos_NCoND_DR1.cufeDocReferenciado = Convert.ToString(OCUFEInvoice.Fields.Item("CUFE").Value.ToString());

                    Datos_NCoND_DR1.descripcion = descripcion;

                    Datos_NCoND_DR1.numeroDocumento = Convert.ToString(OCUFEInvoice.Fields.Item("consecutivoDocumento").Value.ToString());

                    FacturadeVenta.documentosReferenciados[0] = Datos_NCoND_DR1;

                    #endregion

                    #region Documento Referenciado 2

                    DocumentoReferenciado Datos_NCoND_DR2 = new DocumentoReferenciado();

                    Datos_NCoND_DR2.codigoInterno = "5";
                    Datos_NCoND_DR2.cufeDocReferenciado = Convert.ToString(OCUFEInvoice.Fields.Item("CUFE").Value.ToString());
                    Datos_NCoND_DR2.fecha = Convert.ToString(OCUFEInvoice.Fields.Item("fechaEmision").Value.ToString());
                    Datos_NCoND_DR2.numeroDocumento = Convert.ToString(OCUFEInvoice.Fields.Item("consecutivoDocumento").Value.ToString());
                    Datos_NCoND_DR2.tipoCUFE = "CUFE-SHA384";

                    FacturadeVenta.documentosReferenciados[1] = Datos_NCoND_DR2;

                    #endregion

                }


            }

            #endregion

            #region Impuestos

            int CantidadImpuestosGenerales;
            int CantidadImpuestosTotales;
            int SecuenciaArregloImpuestos;
            int PosicionImpuestos;

            CantidadImpuestosGenerales = oImpuestos.RecordCount;

            #region Valida si exiten impuestos y asigna a "ImpuestosGenerales"

            if (CantidadImpuestosGenerales > 0)
            {
                FacturadeVenta.impuestosGenerales = new FacturaImpuestos[CantidadImpuestosGenerales];

                oImpuestos.MoveFirst();

                SecuenciaArregloImpuestos = 0;
                PosicionImpuestos = SecuenciaArregloImpuestos + 1;

                do
                {
                    #region Asignacion impuestosGenerales

                    FacturaImpuestos ImpuestosGenerales = new FacturaImpuestos();

                    #region Detalle impuestosGenerales

                    string sbaseImponibleTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("baseImponibleTOTALImp").Value.ToString());
                    ImpuestosGenerales.baseImponibleTOTALImp = sbaseImponibleTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    ImpuestosGenerales.codigoTOTALImp = Convert.ToString(oImpuestos.Fields.Item("codigoTOTALImp").Value.ToString());

                    string sporcentajeTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("porcentajeTOTALImp").Value.ToString());
                    ImpuestosGenerales.porcentajeTOTALImp = sporcentajeTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    ImpuestosGenerales.unidadMedida = Convert.ToString(oImpuestos.Fields.Item("unidadMedida").Value.ToString());

                    string svalorTOTALImp_ImpuestosGenerales = Convert.ToString(oImpuestos.Fields.Item("valorTOTALImp").Value.ToString());
                    ImpuestosGenerales.valorTOTALImp = svalorTOTALImp_ImpuestosGenerales.Replace(",", ".");

                    #endregion

                    FacturadeVenta.impuestosGenerales[SecuenciaArregloImpuestos] = ImpuestosGenerales;

                    SecuenciaArregloImpuestos++;
                    PosicionImpuestos++;

                    oImpuestos.MoveNext();

                    #endregion

                } while (oImpuestos.EoF == false);
            }

            #endregion

            #region Valida si exiten impuestos y asigna a "impuestosTotales"

            CantidadImpuestosTotales = oImpuestosTotales.RecordCount;

            if (CantidadImpuestosTotales > 0)
            {
                FacturadeVenta.impuestosTotales = new ImpuestosTotales[CantidadImpuestosTotales];

                oImpuestosTotales.MoveFirst();

                SecuenciaArregloImpuestos = 0;
                PosicionImpuestos = SecuenciaArregloImpuestos + 1;

                do
                {
                    #region Asignacion ImpuestosTotales 

                    ImpuestosTotales ImpuestosTotales = new ImpuestosTotales();

                    #region Detalle ImpuestosTotales

                    ImpuestosTotales.codigoTOTALImp = Convert.ToString(oImpuestosTotales.Fields.Item("codigoTOTALImp").Value.ToString());

                    string smontoTotal_ImpuestosTotales = Convert.ToString(oImpuestosTotales.Fields.Item("valorTOTALImp").Value.ToString());
                    ImpuestosTotales.montoTotal = smontoTotal_ImpuestosTotales.Replace(",", ".");

                    #endregion

                    FacturadeVenta.impuestosTotales[SecuenciaArregloImpuestos] = ImpuestosTotales;

                    SecuenciaArregloImpuestos++;
                    PosicionImpuestos++;

                    oImpuestosTotales.MoveNext();

                    #endregion

                } while (oImpuestosTotales.EoF == false);

            }

            #endregion

            #endregion

            #region mediosDePago

            FacturadeVenta.mediosDePago = new MediosDePago[1];

            MediosDePago MediosPago = new MediosDePago();

            MediosPago.medioPago = Convert.ToString(oCabecera.Fields.Item("medioPago").Value.ToString());
            MediosPago.metodoDePago = Convert.ToString(oCabecera.Fields.Item("FormaPago").Value.ToString());

            if (oCabecera.Fields.Item("FormaPago").Value.ToString() == "2")
            {
                MediosPago.fechaDeVencimiento = Convert.ToString(oCabecera.Fields.Item("fechaDeVencimiento").Value.ToString());
            }

            MediosPago.numeroDeReferencia = Convert.ToString(oCabecera.Fields.Item("numeroDeReferencia").Value.ToString());

            FacturadeVenta.mediosDePago[0] = MediosPago;

            #endregion

            #region Informcion Adicional

            if (Convert.ToString(oCabecera.Fields.Item("FacturaTieneMuestras").Value.ToString()) == "SI")
            {
                string[] txtInformacionAdicional = new string[1];

                txtInformacionAdicional[0] = "El total de la Factura a cobrar corresponde a los items registrado sin considerar la muestra gratis";

                FacturadeVenta.informacionAdicional = txtInformacionAdicional;
            }

            #endregion

            return FacturadeVenta;
        }

        public Boolean ExportPDF(SAPbobsCOM.Company _oCompany, string _RutaPDFyXML, string _DocEntry, string _RutaCR, string __TipoDocumento, string _sUserDB, string _sPassDB, string sPathFileLog, string _CreatorUserDoc)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Variables  y objetos

                string sGetRPTDoc = null;
                string sRutaLayout = null;
                string _sMotorDB = null;
                string _sServer = null;
                string _sNameDB = null;
                string _sTipo = null;
                string _UserId = null;
                string _strConnection = null;
                string _sArquitectura = null;

                if (__TipoDocumento == "FacturaDeClientes")
                {
                    _sTipo = "INV2";
                }
                else if (__TipoDocumento == "NotaCreditoClientes")
                {
                    _sTipo = "RIN2";
                }
                else if (__TipoDocumento == "NotaDebitoClientes")
                {
                    _sTipo = "IDN2";
                }

                SAPbobsCOM.Recordset oRGetRPTDoc = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #endregion

                #region Consulta del Motor de Base de datos Y Nombre Base de datos y arquitectura

                _sMotorDB = Convert.ToString(_oCompany.DbServerType);
                _sNameDB = Convert.ToString(_oCompany.CompanyDB);
                _sServer = Convert.ToString(_oCompany.Server);
                _sArquitectura = Convert.ToString(System.IntPtr.Size);

                #endregion

                #region Consulta del nombre del Formato RPT y la ruta donde se encuentra ubicado el RPT

                _UserId = _CreatorUserDoc;

                sGetRPTDoc = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetRPTDocUser");
                sGetRPTDoc = sGetRPTDoc.Replace("%TypeDoc%", _sTipo).Replace("%UserId%", _UserId);

                oRGetRPTDoc.DoQuery(sGetRPTDoc);

                if (oRGetRPTDoc.RecordCount > 0)
                {

                }
                else
                {
                    sGetRPTDoc = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetRPTDoc");
                    sGetRPTDoc = sGetRPTDoc.Replace("%TypeDoc%", _sTipo);

                    oRGetRPTDoc.DoQuery(sGetRPTDoc);
                }

                sRutaLayout = _RutaCR + "\\" + Convert.ToString(oRGetRPTDoc.Fields.Item("NombreFormato").Value.ToString()) + ".rpt";

                #endregion

                #region Generacion del PDF

                if (_sMotorDB == "dst_HANADB")
                {
                    if (_sArquitectura == "8")
                    {
                        #region Genera el PDF con cliente SAP a 64X

                        ReportDocument LayoutPDF = new ReportDocument();

                        LayoutPDF.Load(sRutaLayout);

                        _strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY}", _sUserDB, _sPassDB, _sServer, _sNameDB);

                        NameValuePairs2 logonProps2 = LayoutPDF.DataSourceConnections[0].LogonProperties;
                        logonProps2.Set("Provider", "B1CRHPROXY");
                        logonProps2.Set("Server Type", "B1CRHPROXY");
                        logonProps2.Set("Connection String", _strConnection);

                        LayoutPDF.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        LayoutPDF.DataSourceConnections[0].SetConnection(_sServer, _sNameDB, false);
                        LayoutPDF.SetParameterValue("DocKey@", _DocEntry);
                        LayoutPDF.SetParameterValue("Schema@", _sNameDB);

                        LayoutPDF.ExportToDisk(ExportFormatType.PortableDocFormat, _RutaPDFyXML);

                        LayoutPDF.Close();

                        LayoutPDF.Dispose();

                        GC.SuppressFinalize(LayoutPDF);

                        DllFunciones.Logger("PDF Generado correctamente", sPathFileLog);

                        #endregion
                    }
                    else if (_sArquitectura == "4")
                    {
                        #region Genera el PDF con cliente SAP a 32X

                        ReportDocument LayoutPDF = new ReportDocument();

                        LayoutPDF.Load(sRutaLayout);

                        _strConnection = string.Format("DRIVER={0};UID={1};PWD={2};SERVERNODE={3};DATABASE={4};", "{B1CRHPROXY32}", _sUserDB, _sPassDB, _sServer, _sNameDB);

                        NameValuePairs2 logonProps2 = LayoutPDF.DataSourceConnections[0].LogonProperties;
                        logonProps2.Set("Provider", "B1CRHPROXY32");
                        logonProps2.Set("Server Type", "B1CRHPROXY32");
                        logonProps2.Set("Connection String", _strConnection);

                        LayoutPDF.DataSourceConnections[0].SetLogonProperties(logonProps2);
                        LayoutPDF.DataSourceConnections[0].SetConnection(_sServer, _sNameDB, false);
                        LayoutPDF.SetParameterValue("DocKey@", _DocEntry);
                        LayoutPDF.SetParameterValue("Schema@", _sNameDB);

                        LayoutPDF.ExportToDisk(ExportFormatType.PortableDocFormat, _RutaPDFyXML);

                        LayoutPDF.Close();

                        LayoutPDF.Dispose();

                        GC.SuppressFinalize(LayoutPDF);

                        DllFunciones.Logger("PDF Generado correctamente", sPathFileLog);

                        #endregion
                    }
                }
                else
                {
                    #region Genera el PDF con cliente SAP a 32x o 64x

                    ReportDocument LayoutPDF = new ReportDocument();

                    DiskFileDestinationOptions DestinoDocumento = new DiskFileDestinationOptions();
                    PdfRtfWordFormatOptions OpcionesPDF = new PdfRtfWordFormatOptions();

                    LayoutPDF.Load(sRutaLayout);

                    int Contador = LayoutPDF.DataSourceConnections.Count;
                    LayoutPDF.DataSourceConnections[0].IntegratedSecurity = false;
                    LayoutPDF.DataSourceConnections[0].SetLogon(_sUserDB, _sPassDB);
                    ExportOptions OpExport = LayoutPDF.ExportOptions;
                    OpExport.ExportDestinationType = ExportDestinationType.DiskFile;
                    OpExport.ExportFormatType = ExportFormatType.PortableDocFormat;
                    DestinoDocumento.DiskFileName = _RutaPDFyXML;
                    OpExport.ExportDestinationOptions = (ExportDestinationOptions)DestinoDocumento;
                    OpExport.ExportFormatOptions = (ExportFormatOptions)OpcionesPDF;

                    LayoutPDF.SetParameterValue("DocKey@", _DocEntry);

                    LayoutPDF.Export();

                    LayoutPDF.Close();

                    LayoutPDF.Dispose();

                    GC.SuppressFinalize(LayoutPDF);

                    DllFunciones.Logger("PDF Generado correctamente", sPathFileLog);

                    #endregion
                }

                #endregion

                #region Libreacion de Objetos

                DllFunciones.liberarObjetos(oRGetRPTDoc);

                #endregion

            }
            catch (Exception e)
            {
                DllFunciones.Logger("No se pudo generar el PDF, error: " + e, sPathFileLog);
            }

            return true;
        }

        public void EnviarDocumentosMasivamenteTFHKA(SAPbobsCOM.Company _oCompany, string _TipoDocumento, string TipoIntegracion, string sDocEntryInvoice, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                if (TipoIntegracion == "M")
                {

                }
                else if (TipoIntegracion == "S")
                {
                    #region Envio del documento por el visor de documentos

                    #region Consulta URL

                    string sGetModo = null;
                    string sURLEmision = null;
                    string sURLAdjuntos = null;
                    string sModo = null;

                    SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetModoandURL");

                    sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                    oConsultarGetModo.DoQuery(sGetModo);

                    sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                    sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                    sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());

                    DllFunciones.liberarObjetos(oConsultarGetModo);

                    #endregion

                    #region Instanciacion parametros TFHKA

                    //Especifica el puerto (HTTP o HTTPS)
                    if (sModo == "PRU")
                    {
                        BasicHttpBinding port = new BasicHttpBinding();
                    }
                    else if (sModo == "PRO")
                    {
                        BasicHttpsBinding port = new BasicHttpsBinding();
                    }

                    port.MaxBufferPoolSize = Int32.MaxValue;
                    port.MaxBufferSize = Int32.MaxValue;
                    port.MaxReceivedMessageSize = Int32.MaxValue;
                    port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                    port.SendTimeout = TimeSpan.FromMinutes(2);
                    port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                    if (sModo == "PRO")
                    {
                        port.Security.Mode = BasicHttpSecurityMode.Transport;
                    }

                    //Especifica la dirección de conexion para Emision y Adjuntos 
                    EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION
                    EndpointAddress endPointAdjuntos = new EndpointAddress(sURLAdjuntos); //URL DEMO ADJUNTOS          

                    #endregion

                    #region Variables y objetos

                    string sDocNumInvoice = null;
                    string sQueryDocEntryDocument = null;
                    string sProcedureXML = null;
                    string sDocumentoCabecera = null;
                    string sDocumentoLinea = null;
                    string sDocumentoImpuestosGenerales = null;
                    string sDocumentoImpuestosTotales = null;
                    string sParametrosTFHKA = null;
                    string sRutaCR = null;
                    string sPrefijoConDoc = null;
                    string sPrefijo = null;
                    string sStatusDoc = null;
                    string sFormaEnvio = null;
                    string sLlave = null;
                    string sPassword = null;
                    string sUserDB = null;
                    string sPassDB = null;
                    string sRutaPDF = null;
                    string sRutaXML = null;
                    string sNombreDocumento = null;
                    string sNombreDocWarning = null;
                    string sCUFEInvoice = null;
                    string sGenerarXMLPrueba = null;
                    string CreatorUserDoc = null;

                    Boolean GeneroPDF = false;

                    if (_TipoDocumento == "FacturaDeClientes")
                    {
                        sNombreDocumento = "Factura_de_Venta_No_";
                        sNombreDocWarning = "Factura de venta";
                    }
                    else if (_TipoDocumento == "NotaCreditoClientes")
                    {
                        sNombreDocumento = "Nota_Credito_No_";
                        sNombreDocWarning = "Nota credito de clientes";
                    }
                    else if (_TipoDocumento == "NotaDebitoClientes")
                    {
                        sNombreDocumento = "Nota_debito_Clientes_No_";
                        sNombreDocWarning = "Nota debito de clientes";
                    }

                    #endregion

                    #region Consulta de documento en la base de datos y el estado del documento

                    SAPbobsCOM.Recordset oConsultaDocEntry = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    sQueryDocEntryDocument = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetDocEntryAndParametersService");

                    if (_TipoDocumento == "FacturaDeClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "--");
                    }
                    else if (_TipoDocumento == "NotaCreditoClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "ORIN").Replace("%DocSubType%", "--");
                    }
                    else if (_TipoDocumento == "NotaDebitoClientes")
                    {
                        sQueryDocEntryDocument = sQueryDocEntryDocument.Replace("%sDocEntry%", sDocEntryInvoice).Replace("%Tabla%", "OINV").Replace("%DocSubType%", "DN");
                    }

                    oConsultaDocEntry.DoQuery(sQueryDocEntryDocument);

                    sDocNumInvoice = Convert.ToString(oConsultaDocEntry.Fields.Item("DocNum").Value.ToString());
                    sPrefijo = Convert.ToString(oConsultaDocEntry.Fields.Item("PrefijoDes").Value.ToString());
                    sStatusDoc = Convert.ToString(oConsultaDocEntry.Fields.Item("CRWS").Value.ToString());
                    sFormaEnvio = Convert.ToString(oConsultaDocEntry.Fields.Item("FormaEnvio").Value.ToString());
                    sLlave = Convert.ToString(oConsultaDocEntry.Fields.Item("Llave").Value.ToString());
                    sPassword = Convert.ToString(oConsultaDocEntry.Fields.Item("Password").Value.ToString());
                    sUserDB = Convert.ToString(oConsultaDocEntry.Fields.Item("UserDB").Value.ToString());
                    sPassDB = Convert.ToString(oConsultaDocEntry.Fields.Item("PassDB").Value.ToString());
                    sRutaXML = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaXML").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".txt";
                    sRutaPDF = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaPDF").Value.ToString()) + "\\" + sNombreDocumento + sPrefijo + '_' + sDocNumInvoice + ".pdf";
                    sRutaCR = Convert.ToString(oConsultaDocEntry.Fields.Item("RutaCR").Value.ToString());
                    sGenerarXMLPrueba = Convert.ToString(oConsultaDocEntry.Fields.Item("GeneraXMLP").Value.ToString());
                    CreatorUserDoc = Convert.ToString(oConsultaDocEntry.Fields.Item("UserSign").Value.ToString());

                    #endregion

                    if (sStatusDoc == "200")
                    {
                        DllFunciones.liberarObjetos(oConsultaDocEntry);
                        DllFunciones.liberarObjetos(oConsultarGetModo);
                    }
                    else
                    {
                        DllFunciones.Logger(" Enviando " + sNombreDocWarning + " No. " + sDocNumInvoice, sPathFileLog);

                        if (oConsultaDocEntry.RecordCount > 0)
                        {

                            #region Si existe el numero de factura, busca la factura y crea el objeto factura

                            SAPbobsCOM.Recordset oCabeceraDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oLineasDocumento = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oImpuestosGenerales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oImpuestosTotales = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            SAPbobsCOM.Recordset oCUFEInvoice = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sProcedureXML = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "ExecProcedureBOFacturaXML");

                            if (_TipoDocumento == "FacturaDeClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }
                            else if (_TipoDocumento == "NotaCreditoClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "14").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }
                            else if (_TipoDocumento == "NotaDebitoClientes")
                            {
                                sDocumentoCabecera = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Encabezado");
                                sDocumentoLinea = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Lineas");
                                sDocumentoImpuestosGenerales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "Impuestos");
                                sDocumentoImpuestosTotales = sProcedureXML.Replace("%DocEntry%", sDocEntryInvoice).Replace("%ObjecType%", "13_ND").Replace("%TipoConsulta%", "ImpuestosTotales");
                            }

                            oCabeceraDocumento.DoQuery(sDocumentoCabecera);
                            oLineasDocumento.DoQuery(sDocumentoLinea);
                            oImpuestosGenerales.DoQuery(sDocumentoImpuestosGenerales);
                            oImpuestosTotales.DoQuery(sDocumentoImpuestosTotales);

                            if (_TipoDocumento == "NotaCreditoClientes")
                            {
                                sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetCUFEInvoice");
                                sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                oCUFEInvoice.DoQuery(sCUFEInvoice);
                            }
                            else if (_TipoDocumento == "NotaDebitoClientes")
                            {
                                sCUFEInvoice = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetCUFEDebitNote");
                                sCUFEInvoice = sCUFEInvoice.Replace("%DocNum%", Convert.ToString(oCabeceraDocumento.Fields.Item("No_FV").Value.ToString()));

                                oCUFEInvoice.DoQuery(sCUFEInvoice);

                            }

                            FacturaGeneral Documento = oBuillInvoice(oCabeceraDocumento, oLineasDocumento, oImpuestosGenerales, oImpuestosTotales, oCUFEInvoice, _TipoDocumento);

                            #endregion

                            #region Guarda el TXT en la ruta del XML configurada

                            StreamWriter MyFile = new StreamWriter(sRutaXML); //ruta y name del archivo request a almecenar

                            #endregion

                            #region Serealizando el documento

                            SAPbobsCOM.Recordset oParametrosTFHKA = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            sParametrosTFHKA = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetParameterstoSend");

                            oParametrosTFHKA.DoQuery(sParametrosTFHKA);

                            XmlSerializer Serializer1 = new XmlSerializer(typeof(FacturaGeneral));
                            Serializer1.Serialize(MyFile, Documento); // Objeto serializado
                            MyFile.Close();

                            if (sGenerarXMLPrueba == "N")
                            {
                                File.Delete(sRutaXML);
                            }

                            #endregion

                            #region Envio del objeto factura a TFHKA

                            serviceClient = new BOeBillingService.ServicioEmisionFE.ServiceClient(port, endPointEmision);
                            serviceClientAdjuntos = new BOeBillingService.ServicioAdjuntosFE.ServiceClient(port, endPointAdjuntos);

                            DocumentResponse RespuestaDoc = new BOeBillingService.ServicioEmisionFE.DocumentResponse(); //objeto Response del metodo enviar

                            if (string.IsNullOrEmpty(sLlave))
                            {
                                DllFunciones.Logger("Error Paso 4: No se ha parametrizado la llave de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN ", sPathFileLog);
                            }
                            else if (string.IsNullOrEmpty(sPassword))
                            {
                                DllFunciones.Logger("Error Paso 4: No se ha parametrizado el password de TFHKA en la configuracion Inicial, por lo cual no se puede enviar la factura a la DIAN", sPathFileLog);
                            }
                            else
                            {
                                #region Respuesta el Web Service de TFHKA y actualizacion de los campos en la factura

                                RespuestaDoc = serviceClient.Enviar(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), Documento, sFormaEnvio);

                                if (RespuestaDoc.codigo == 200)
                                {
                                    #region Procesa la repuesta

                                    DllFunciones.Logger("Se envio correctamente la " + sNombreDocWarning + "No. " + sDocNumInvoice, sPathFileLog);

                                    #region Se actualiza el documento en SAP con las respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, "Documento autorizado por la DIAN", RespuestaDoc.cufe, RespuestaDoc.qr, null, null, sPathFileLog);
                                    }

                                    #endregion

                                    #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {

                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB, sPathFileLog, CreatorUserDoc);
                                        }
                                    }

                                    #endregion

                                    #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                    if (GeneroPDF == true)
                                    {
                                        //DllFunciones.StatusBar(_sboapp, BoStatusBarMessageType.smt_Success, "Paso 7: Adjuntando PDF, por favor espere ...");

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null, sPathFileLog);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null, sPathFileLog);
                                        }

                                    }
                                    else
                                    {
                                    }

                                    #endregion

                                    #region Envia el PDF al proveedor tecnologico TFHKA

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;

                                    EnviarAdjuntosTFHKA(_oCompany, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPathFileLog, "Normal");

                                    #endregion

                                    #region Se descarga el XML y se adjunta a la factura de venta   

                                    #region Descarga el XML y retorna la confirmacion

                                    bool DescargoXML = false;

                                    DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML, sPathFileLog);

                                    #endregion

                                    #region Actualiza el campo de XML en el documento de SAP

                                    if (DescargoXML == true)
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"), sPathFileLog);
                                    }
                                    else
                                    {

                                    }

                                    #endregion

                                    #endregion

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 201)
                                {
                                    #region Procesa la respuesta                                                        

                                    #region Consulta el estado del documento en TFHKA

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                    DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                    #endregion

                                    #region Actualiza el documento con la respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null, sPathFileLog);
                                    }
                                    #endregion

                                    if (resp.codigo == 200)
                                    {
                                        #region Valida la forma de envio,si es 11,  genera el PDF y retorna confirmacion de la generacion del PDF

                                        if (sFormaEnvio == "11")
                                        {
                                            FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                            if (ValidacionPDF.Exists)
                                            {
                                                GeneroPDF = true;
                                            }
                                            else
                                            {
                                                GeneroPDF = ExportPDF(_oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB, sPathFileLog, CreatorUserDoc);
                                            }

                                        }
                                        else
                                        {

                                        }

                                        #endregion
                                    }

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 101)
                                {
                                    #region Procesa la respuesta 

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + Convert.ToString(RespuestaDoc.mensaje) + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje, "", "", null, null, sPathFileLog);
                                    }

                                    #endregion

                                }
                                else if (RespuestaDoc.codigo == 99)
                                {

                                    #region Procesa la respuesta

                                    string sRsErrorDIAN99 = null;

                                    for (int i = 0; i < RespuestaDoc.reglasValidacionDIAN.Length; i++)
                                    {
                                        sRsErrorDIAN99 = sRsErrorDIAN99 + " ; " + Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(i));
                                    }

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + sRsErrorDIAN99 + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.reglasValidacionDIAN.GetValue(0)), "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null, sPathFileLog);
                                    }

                                    #endregion
                                }

                                else if (RespuestaDoc.codigo == 109)
                                {
                                    #region Procesa la respuesta

                                    string sRsErrorDIAN109 = null;

                                    for (int i = 0; i < RespuestaDoc.mensajesValidacion.Length; i++)
                                    {
                                        sRsErrorDIAN109 = sRsErrorDIAN109 + " ; " + Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(i));
                                    }

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + sRsErrorDIAN109 + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, Convert.ToString(RespuestaDoc.mensajesValidacion.GetValue(0)), "", "", null, null, sPathFileLog);
                                    }

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 110)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + Convert.ToString(RespuestaDoc.mensaje.ToString()) + " Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos  ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes")) if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null, sPathFileLog);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString() + ", Total de Factura es diferente de la suma de Total valor bruto + Tributos - Total Tributo Retenidos - Anticipos ", "", "", null, null, sPathFileLog);
                                        }


                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 111)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + Convert.ToString(RespuestaDoc.mensaje.ToString()) + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }

                                    #endregion

                                }
                                else if (RespuestaDoc.codigo == 112)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + Convert.ToString(RespuestaDoc.mensaje.ToString()) + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 150)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.Logger("Error : Codigo: " + Convert.ToString(RespuestaDoc.codigo) + " : " + Convert.ToString(RespuestaDoc.mensaje.ToString()) + " ", sPathFileLog);

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, RespuestaDoc.codigo, RespuestaDoc.mensaje.ToString(), "", "", null, null, sPathFileLog);
                                    }

                                    #endregion
                                }
                                else if (RespuestaDoc.codigo == 114)
                                {
                                    #region Procesa la respuesta

                                    DllFunciones.Logger("Se envio correctamente la " + sNombreDocWarning + "No. " + sDocNumInvoice, sPathFileLog);

                                    #region Consulta el estado del documento en el proveedor tecnologico

                                    sPrefijoConDoc = Convert.ToString(oCabeceraDocumento.Fields.Item("Prefijo").Value.ToString()) + sDocNumInvoice;
                                    DocumentStatusResponse resp = serviceClient.EstadoDocumento(Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sPrefijoConDoc);

                                    #endregion

                                    #region Se actualiza la factura con las respuesta de TFHKA

                                    if (_TipoDocumento == "FacturaDeClientes" || (_TipoDocumento == "NotaDebitoClientes"))
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null, sPathFileLog);
                                    }
                                    else if (_TipoDocumento == "NotaCreditoClientes")
                                    {
                                        UpdateoCreditNote(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, null, null, sPathFileLog);
                                    }

                                    #endregion

                                    #region Valida la forma de envio, si es 11 genera el PDF y retorna confirmacion de la generacion del PDF

                                    if (sFormaEnvio == "11")
                                    {
                                        FileInfo ValidacionPDF = new FileInfo(sRutaPDF);

                                        if (ValidacionPDF.Exists)
                                        {
                                            GeneroPDF = true;
                                        }
                                        else
                                        {
                                            GeneroPDF = ExportPDF(_oCompany, sRutaPDF, sDocEntryInvoice, sRutaCR, _TipoDocumento, sUserDB, sPassDB, sPathFileLog, CreatorUserDoc);
                                        }

                                    }

                                    #endregion

                                    #region Si genera correctamente el PDF lo adjunta a la factura de venta en SAP, 

                                    if (GeneroPDF == true)
                                    {
                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null, sPathFileLog);
                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, sDocEntryInvoice, resp.codigo, "Documento autorizado por la DIAN", resp.cufe, resp.cadenaCodigoQR, sRutaPDF, null, sPathFileLog);
                                        }

                                        #region Envia el PDF al proveedor tecnologico TFHKA

                                        EnviarAdjuntosTFHKA(_oCompany, oCabeceraDocumento, sRutaPDF, sPrefijoConDoc, sLlave, sPassword, sPathFileLog, "Normal");

                                        #endregion

                                    }
                                    else
                                    {
                                        UpdateoInvoice(_oCompany, sDocEntryInvoice, 0, null, null, null, null, null, sPathFileLog);
                                    }

                                    #endregion

                                    #region Se descarga el XML y se adjunta a la factura de venta

                                    #region Descarga el XML y retorna la confirmacion

                                    bool DescargoXML = false;

                                    DescargoXML = DescargaXML(_oCompany, sPrefijoConDoc, Convert.ToString(oParametrosTFHKA.Fields.Item("TokenEmpresa").Value.ToString()), Convert.ToString(oParametrosTFHKA.Fields.Item("TokenPassword").Value.ToString()), sRutaXML, sPathFileLog);

                                    #endregion

                                    #region Actualiza el campo de XML en el documento de SAP

                                    if (DescargoXML == true)
                                    {

                                        if (_TipoDocumento == "FacturaDeClientes")
                                        {
                                            UpdateoInvoice(_oCompany, sDocEntryInvoice, 0, null, null, null, null, sRutaXML.Replace(".txt", ".xml"), sPathFileLog);

                                        }
                                        else if (_TipoDocumento == "NotaCreditoClientes")
                                        {
                                            UpdateoCreditNote(_oCompany, sDocEntryInvoice, 0, null, null, null, sRutaPDF, null, sPathFileLog);
                                        }

                                    }
                                    else
                                    {

                                    }

                                    #endregion

                                    #endregion

                                    #endregion
                                }

                                #endregion
                            }

                            #endregion
                        }
                        else
                        {

                        }
                    }
                    #endregion
                }
            }
            catch (Exception ex)
            {
                DllFunciones.Logger(" No se pudo enviar la " + _TipoDocumento + " con DocEntry" + sDocEntryInvoice, sPathFileLog);
                DllFunciones.Logger(ex.Message.ToString(), sPathFileLog);
                DllFunciones.Logger(ex.StackTrace.ToString(), sPathFileLog);

            }
        }

        public void EnviarAdjuntosTFHKA(SAPbobsCOM.Company _oCompany, SAPbobsCOM.Recordset oCabecera, string _RutaPDFyXML, string _sPrefijoConDoc, string _tbxTokenEmpresa, string _tbxTokenPassword, string sPathFileLog, string _sMetodo)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            for (int i = 0; i < 1; i++)
            {
                FileInfo file = new FileInfo(_RutaPDFyXML);

                if (file.Exists)
                {
                    BinaryReader bReader = new BinaryReader(file.OpenRead());
                    byte[] anexByte = bReader.ReadBytes((int)file.Length);
                    //anexB64 = Convert.ToBase64String(anexByte);
                    ServicioAdjuntosFE.CargarAdjuntos uploadAttachment = new ServicioAdjuntosFE.CargarAdjuntos();
                    uploadAttachment.archivo = anexByte;
                    uploadAttachment.numeroDocumento = _sPrefijoConDoc;

                    #region Revision Correos a Enviar

                    #region Variables Correo

                    string CorreoDeEntrega1 = Convert.ToString(oCabecera.Fields.Item("correoEntrega1").Value.ToString());
                    string CorreoDeEntrega2 = Convert.ToString(oCabecera.Fields.Item("correoEntrega2").Value.ToString());
                    string CorreoDeEntrega3 = Convert.ToString(oCabecera.Fields.Item("correoEntrega3").Value.ToString());
                    string CorreoDeEntrega4 = Convert.ToString(oCabecera.Fields.Item("correoEntrega4").Value.ToString());
                    string CorreoDeEntrega5 = Convert.ToString(oCabecera.Fields.Item("correoEntrega5").Value.ToString());

                    int ContadorCorreos = 0;

                    #endregion

                    #region Contador de los correos a enviar 

                    if (string.IsNullOrEmpty(CorreoDeEntrega1))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega2))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega3))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega4))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    if (string.IsNullOrEmpty(CorreoDeEntrega5))
                    {

                    }
                    else
                    {
                        ContadorCorreos++;
                    }

                    #endregion

                    string[] correoEntrega = new string[ContadorCorreos];

                    #region Asignacion de los correos a enviar

                    if (ContadorCorreos == 1)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                    }
                    else if (ContadorCorreos == 2)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                    }
                    else if (ContadorCorreos == 3)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                    }
                    else if (ContadorCorreos == 4)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                        correoEntrega[3] = CorreoDeEntrega4;
                    }
                    else if (ContadorCorreos == 5)
                    {
                        correoEntrega[0] = CorreoDeEntrega1;
                        correoEntrega[1] = CorreoDeEntrega2;
                        correoEntrega[2] = CorreoDeEntrega3;
                        correoEntrega[3] = CorreoDeEntrega4;
                        correoEntrega[4] = CorreoDeEntrega5;
                    }

                    #endregion

                    #endregion

                    uploadAttachment.email = correoEntrega;
                    uploadAttachment.nombre = file.Name.Substring(0, file.Name.Length - 4);
                    uploadAttachment.formato = file.Extension.Substring(1);
                    uploadAttachment.tipo = "1";
       
                    if (Convert.ToString(oCabecera.Fields.Item("notificar").Value.ToString()) == "NO")
                    {
                        uploadAttachment.enviar = "0";
                    }
                    else
                    {
                        if (i + 1 == 1)
                        {
                            uploadAttachment.enviar = "1";
                        }
                        else
                        {
                            uploadAttachment.enviar = "0";
                        }
                    }

                    ServicioAdjuntosFE.UploadAttachmentResponse fileRespuesta = new ServicioAdjuntosFE.UploadAttachmentResponse();
                    fileRespuesta = serviceClientAdjuntos.CargarAdjuntos(_tbxTokenEmpresa, _tbxTokenPassword, uploadAttachment);
                    if (fileRespuesta.codigo == 200)
                    {
                        if (_sMetodo == "Validacion")
                        {
                            InsertSendEmail(_oCompany, oCabecera, null, sPathFileLog, Convert.ToString(fileRespuesta.codigo));
                        }
                    }
                    else
                    {
                        DllFunciones.Logger("No se pudo cargar el PDF: " + fileRespuesta.mensaje, sPathFileLog);

                        if (_sMetodo == "Validacion")
                        {
                            InsertSendEmail(_oCompany, oCabecera, null, sPathFileLog, Convert.ToString(fileRespuesta.codigo));
                        }
                    }
                }
                else
                {
                    DllFunciones.Logger("No se pudo enviar el PDF, no existe el archivo para el documento " + _sPrefijoConDoc, sPathFileLog);
                }

            }

        }

        public void UpdateEmailSendInSAP(SAPbobsCOM.Company _oCompany, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;
                string sllave = null;
                string sPass = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());
                sllave = Convert.ToString(oConsultarGetModo.Fields.Item("Llave").Value.ToString());
                sPass = Convert.ToString(oConsultarGetModo.Fields.Item("Password").Value.ToString());

                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)

                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();

                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                if (sModo == "PRO")
                {
                    port.Security.Mode = BasicHttpSecurityMode.Transport;
                }

                //Especifica la dirección de conexion para Emision y Adjuntos 
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION
                EndpointAddress endPointAdjuntos = new EndpointAddress(sURLAdjuntos); //URL DEMO ADJUNTOS    

                serviceClient = new ServicioEmisionFE.ServiceClient(port, endPointEmision);
                serviceClientAdjuntos = new ServicioAdjuntosFE.ServiceClient(port, endPointAdjuntos);

                DocumentStatusResponse resp = new BOeBillingService.ServicioEmisionFE.DocumentStatusResponse();

                #endregion

                #region Variables y Objetos 

                SAPbobsCOM.Recordset oGetDocs = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oGetDocsPDF = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string sGetDocs = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetEmailNotSend");

                sGetDocs = sGetDocs.Replace("%Table%", "OINV");

                oGetDocs.DoQuery(sGetDocs);

                int iDocumentosaconsultar = oGetDocs.RecordCount;

                #endregion

                #region Consulta si ya se enviaron los correos en Facturas de clientes

                DllFunciones.Logger("Sincronizando estado envio correos facturas de venta", sPathFileLog);

                if (iDocumentosaconsultar > 0)
                {
                    oGetDocs.MoveFirst();

                    do
                    {
                        string sDocPrefijo = null;
                        sDocPrefijo = Convert.ToString(oGetDocs.Fields.Item("DocPref").Value.ToString());

                        resp = serviceClient.EstadoDocumento(sllave, sPass, sDocPrefijo);

                        if (resp.codigo == 200)
                        {
                            InsertSendEmail(_oCompany, oGetDocs, resp, sPathFileLog, null);
                            DllFunciones.Logger("Documento No." + sDocPrefijo + " sincronizado", sPathFileLog);
                        }

                        oGetDocs.MoveNext();

                    } while (oGetDocs.EoF == false);

                    DllFunciones.Logger("Sincronizacion finalizada de correos de facturas", sPathFileLog);
                }
                else
                {
                    DllFunciones.Logger("No se encontraron correos de facturas de venta a sincronizar", sPathFileLog);
                }

                #endregion

                #region Enviar PDF Pendientes

                DllFunciones.Logger("Sincronizando PDF pendientes Notas Credito ", sPathFileLog);

                oGetDocsPDF.DoQuery(sGetDocs);

                int iDocumentosaconsultarPDF = oGetDocsPDF.RecordCount;

                if (iDocumentosaconsultarPDF > 0)
                {
                    oGetDocsPDF.MoveFirst();
                    
                    do
                    {

                        if (Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_PdfTFHKA").Value.ToString()) != "200")
                        {

                            if (Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_PdfTFHKA").Value.ToString()) != "107")
                            {

                                string sDocPrefijo = null;
                                sDocPrefijo = Convert.ToString(oGetDocsPDF.Fields.Item("DocPref").Value.ToString());

                                string sRutaPDF = null;
                                sRutaPDF = Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_RPDF").Value.ToString());

                                FileInfo file = new FileInfo(sRutaPDF);

                                if (file.Exists)
                                {
                                    EnviarAdjuntosTFHKA(_oCompany, oGetDocsPDF, sRutaPDF, sDocPrefijo, sllave, sPass, sPathFileLog, "Validacion");
                                }
                                else
                                {
                                    InsertSendEmail(_oCompany, oGetDocsPDF, null, sPathFileLog, "99");
                                }

                                oGetDocsPDF.MoveNext();

                            }

                        }

                    } while (oGetDocsPDF.EoF == false);

                }
                else
                {
                    DllFunciones.Logger("No se encontraron PDF por pendientes por sincronizar ", sPathFileLog);
                }

                #endregion

                #region Consulta si ya se enviaron los correos en notas credito de clientes

                DllFunciones.Logger("Sincronizando estado envio correos notas credito", sPathFileLog);

                sGetDocs = null;

                sGetDocs = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetEmailNotSend");

                sGetDocs = sGetDocs.Replace("%Table%", "ORIN").Replace("FVC","NCC");

                oGetDocs.DoQuery(sGetDocs);

                int iDocumentosaconsultarORIN = oGetDocs.RecordCount;

                if (iDocumentosaconsultarORIN > 0)
                {
                    oGetDocs.MoveFirst();

                    do
                    {
                        string sDocPrefijo = null;
                        sDocPrefijo = Convert.ToString(oGetDocs.Fields.Item("DocPref").Value.ToString());

                        resp = serviceClient.EstadoDocumento(sllave, sPass, sDocPrefijo);

                        if (resp.codigo == 200)
                        {
                            InsertSendEmail(_oCompany, oGetDocs, resp, sPathFileLog, null);
                            DllFunciones.Logger("Documento No." + sDocPrefijo + " sincronizado", sPathFileLog);
                        }

                        oGetDocs.MoveNext();

                    } while (oGetDocs.EoF == false);

                    DllFunciones.Logger("Sincronizacion finalizada", sPathFileLog);
                }
                else
                {
                    DllFunciones.Logger("No se encontraron correos de notas credito a sincronizar", sPathFileLog);
                }

                #endregion

                #region Enviar PDF Pendientes

                DllFunciones.Logger("Sincronizando PDF pendientes", sPathFileLog);

                oGetDocsPDF.DoQuery(sGetDocs);

                int iDocumentosaconsultarPDFORIN = oGetDocsPDF.RecordCount;

                if (iDocumentosaconsultarPDFORIN > 0)
                {
                    oGetDocsPDF.MoveFirst();

                    do
                    {

                        if (Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_PdfTFHKA").Value.ToString()) != "200")
                        {

                            if (Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_PdfTFHKA").Value.ToString()) != "107")
                            {

                                string sDocPrefijo = null;
                                sDocPrefijo = Convert.ToString(oGetDocsPDF.Fields.Item("DocPref").Value.ToString());

                                string sRutaPDF = null;
                                sRutaPDF = Convert.ToString(oGetDocsPDF.Fields.Item("U_BO_RPDF").Value.ToString());

                                FileInfo file = new FileInfo(sRutaPDF);

                                if (file.Exists)
                                {
                                    EnviarAdjuntosTFHKA(_oCompany, oGetDocsPDF, sRutaPDF, sDocPrefijo, sllave, sPass, sPathFileLog, "Validacion");
                                }
                                else
                                {
                                    InsertSendEmail(_oCompany, oGetDocsPDF, null, sPathFileLog, "99");
                                }

                                oGetDocsPDF.MoveNext();

                            }

                        }

                    } while (oGetDocsPDF.EoF == false);

                }
                else
                {
                    DllFunciones.Logger("No se encontraron PDF por pendientes por sincronizar ", sPathFileLog);
                }

                #endregion

                #region Liberacion de Objetos

                DllFunciones.liberarObjetos(oGetDocs);

                #endregion

            }
            catch (Exception ex)
            {

                DllFunciones.Logger(ex.ToString(), sPathFileLog);
            }

        }

        private void UpdateoInvoice(SAPbobsCOM.Company __oCompany, string _sQueryDocEntryInvoice, int _CRWS, string _MRWS, string _WSCUFE, string _WSQR, string _RutaPDF, string _RutaXML, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {
                int _DocEntry = Convert.ToInt32(_sQueryDocEntryInvoice);
                Rsd = 0;

                SAPbobsCOM.Documents oInvoice = (SAPbobsCOM.Documents)(__oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices));

                oInvoice.GetByKey(_DocEntry);

                #region Campo CRWS

                if (_CRWS == 0)
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_CRWS").Value = Convert.ToString(_CRWS);
                }

                #endregion

                #region Campo MRWS

                if (string.IsNullOrEmpty(_MRWS))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_MRWS").Value = Convert.ToString(_MRWS);
                }

                #endregion

                oInvoice.UserFields.Fields.Item("U_BO_S").Value = "3";
                oInvoice.UserFields.Fields.Item("U_BO_PP").Value = "M";

                #region Campo WSCUFE

                if (string.IsNullOrEmpty(_WSCUFE))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_CUFE").Value = Convert.ToString(_WSCUFE);
                }

                #endregion

                #region Campo WSQR

                if (string.IsNullOrEmpty(_WSQR))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_QR").Value = Convert.ToString(_WSQR);
                }

                #endregion

                #region Campo RutaPDF

                if (string.IsNullOrEmpty(_RutaPDF))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_RPDF").Value = Convert.ToString(_RutaPDF);
                }

                #endregion

                #region Campo RutaXML

                if (string.IsNullOrEmpty(_RutaXML))
                {
                }
                else
                {
                    oInvoice.UserFields.Fields.Item("U_BO_XML").Value = Convert.ToString(_RutaXML);
                }

                #endregion

                Rsd = oInvoice.Update();

                if (Rsd == 0)
                {
                    DllFunciones.liberarObjetos(oInvoice);
                }
                else
                {
                    DllFunciones.Logger(__oCompany.GetLastErrorDescription(), sPathFileLog);
                }
            }
            catch (Exception)
            {
                DllFunciones.Logger("No se pudo actualizar la factura de venta - Respuesta 200", sPathFileLog);
            }

        }

        private void UpdateoCreditNote(SAPbobsCOM.Company __oCompany, string _sQueryDocEntryInvoice, int _CRWS, string _MRWS, string _WSCUFE, string _WSQR, string _RutaPDF, string _RutaXML, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();


            try
            {
                int _DocEntry = Convert.ToInt32(_sQueryDocEntryInvoice);
                Rsd = 0;

                SAPbobsCOM.Documents oCreditNote = (SAPbobsCOM.Documents)(__oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes));

                oCreditNote.GetByKey(_DocEntry);

                #region Campo CRWS

                if (_CRWS == 0)
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_CRWS").Value = Convert.ToString(_CRWS);
                }

                #endregion

                #region Campo MRWS

                if (string.IsNullOrEmpty(_MRWS))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_MRWS").Value = Convert.ToString(_MRWS);
                }

                #endregion

                oCreditNote.UserFields.Fields.Item("U_BO_S").Value = "3";
                oCreditNote.UserFields.Fields.Item("U_BO_PP").Value = "M";

                #region Campo WSCUFE

                if (string.IsNullOrEmpty(_WSCUFE))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_CUFE").Value = Convert.ToString(_WSCUFE);
                }

                #endregion

                #region Campo WSQR

                if (string.IsNullOrEmpty(_WSQR))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_QR").Value = Convert.ToString(_WSQR);
                }

                #endregion

                #region Campo RutaPDF

                if (string.IsNullOrEmpty(_RutaPDF))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_RPDF").Value = Convert.ToString(_RutaPDF);
                }

                #endregion

                #region Campo RutaXML

                if (string.IsNullOrEmpty(_RutaXML))
                {
                }
                else
                {
                    oCreditNote.UserFields.Fields.Item("U_BO_XML").Value = Convert.ToString(_RutaXML);
                }

                #endregion

                Rsd = oCreditNote.Update();

                if (Rsd == 0)
                {
                    DllFunciones.liberarObjetos(oCreditNote);
                }
                else
                {
                    DllFunciones.Logger(__oCompany.GetLastErrorDescription(), sPathFileLog);
                }
            }
            catch (Exception)
            {
                DllFunciones.Logger("No se pudo actualizar la nota credito de venta", sPathFileLog);
            }

        }

        private bool DescargaXML(SAPbobsCOM.Company _oCompany, string _sPrefijoConDoc, string _tbxTokenEmpresa, string _tbxTokenPassword, string _sRutaXML, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();


            try
            {
                #region Consulta URL

                string sGetModo = null;
                string sURLEmision = null;
                string sURLAdjuntos = null;
                string sModo = null;
                string sRutaXML = null;

                SAPbobsCOM.Recordset oConsultarGetModo = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                sGetModo = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "GetModoandURL");

                sGetModo = sGetModo.Replace("%Estado%", "\"U_BO_Status\" = 'Y'").Replace("%DocEntry%", " ");

                oConsultarGetModo.DoQuery(sGetModo);

                sURLEmision = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/v1.0/Service.svc?wsdl";
                sURLAdjuntos = Convert.ToString(oConsultarGetModo.Fields.Item("URLTFHKA").Value.ToString()) + "/ws/adjuntos/Service.svc?wsdl";
                sModo = Convert.ToString(oConsultarGetModo.Fields.Item("Modo").Value.ToString());
                sRutaXML = _sRutaXML.Replace(".txt", ".xml");


                DllFunciones.liberarObjetos(oConsultarGetModo);

                #endregion

                #region Instanciacion parametros TFHKA

                //Especifica el puerto (HTTP o HTTPS)
                if (sModo == "PRU")
                {
                    BasicHttpBinding port = new BasicHttpBinding();
                }
                else if (sModo == "PRO")
                {
                    BasicHttpsBinding port = new BasicHttpsBinding();
                }

                port.MaxBufferPoolSize = Int32.MaxValue;
                port.MaxBufferSize = Int32.MaxValue;
                port.MaxReceivedMessageSize = Int32.MaxValue;
                port.ReaderQuotas.MaxStringContentLength = Int32.MaxValue;
                port.SendTimeout = TimeSpan.FromMinutes(2);
                port.ReceiveTimeout = TimeSpan.FromMinutes(2);

                //Especifica la dirección de conexion para Demo y Adjuntos para pruebas
                EndpointAddress endPointEmision = new EndpointAddress(sURLEmision); //URL DEMO EMISION

                ServicioEmisionFE.ServiceClient serviceClienTFHKA;

                serviceClienTFHKA = new ServicioEmisionFE.ServiceClient(port, endPointEmision);

                #endregion

                DownloadXMLResponse xmlResponse;

                xmlResponse = serviceClient.DescargaXML(_tbxTokenEmpresa, _tbxTokenPassword, _sPrefijoConDoc);

                if (xmlResponse.codigo == 200)
                {
                    File.WriteAllBytes(sRutaXML, Convert.FromBase64String(xmlResponse.documento));
                    DllFunciones.Logger("XML Descargado correctamente", sPathFileLog);
                    return true;
                }
                else
                {
                    return false;
                }


            }
            catch (Exception)
            {
                DllFunciones.Logger("No se pudo descargar el XML", sPathFileLog);
                return false;
            }

        }

        public void EnviarDocumentosDIANServicioLocalizacion(SAPbobsCOM.Company oCompany, string sPathFileLog)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();

            try
            {

                #region Variables y Objetos

                int DocsCounter = 0;
                string sDocEntry = null;

                string sQryInvoice = null;
                string sQryCreditMemo = null;
                string sQryDebitMemo = null;
                string sQrySeriesNumber = null;
                string _sMotorDB = null;

                SAPbobsCOM.Recordset oRsSeriesNumber = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRsInvoice = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRsCreditMemo = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRsDebitMemo = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _sMotorDB = Convert.ToString(oCompany.DbServerType);

                sQrySeriesNumber = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetSeriesNumberActive");

                oRsSeriesNumber.DoQuery(sQrySeriesNumber);

                if (_sMotorDB == "dst_HANADB")
                {
                    sQryInvoice = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetInvoices");
                    sQryInvoice = sQryInvoice.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND IFNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(0).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                    sQryCreditMemo = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetCreditMemo");
                    sQryCreditMemo = sQryCreditMemo.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND IFNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(1).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                    sQryDebitMemo = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetDebitMemo");
                    sQryDebitMemo = sQryDebitMemo.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND IFNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(2).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                }
                else
                {
                    sQryInvoice = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetInvoices");
                    sQryInvoice = sQryInvoice.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND ISNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(0).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                    sQryCreditMemo = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetCreditMemo");
                    sQryCreditMemo = sQryCreditMemo.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND ISNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(1).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                    sQryDebitMemo = DllFunciones.GetStringXMLDocument(oCompany, "BOeBillingService", "eBilling", "GetDebitMemo");
                    sQryDebitMemo = sQryDebitMemo.Replace("%FI%", "20200101").Replace("%FF%", "20251231").Replace("%EstadoDocumento%", "AND ISNULL(\"U_BO_CRWS\",'') != ('200')").Replace("%Series%", Convert.ToString(oRsSeriesNumber.Fields.Item(2).Value)).Replace("***SN***", " ").Replace("***DocNum***", " ");

                }


                oRsInvoice.DoQuery(sQryInvoice);
                oRsCreditMemo.DoQuery(sQryCreditMemo);
                oRsDebitMemo.DoQuery(sQryDebitMemo);

                #endregion

                DocsCounter = oRsInvoice.RecordCount + oRsCreditMemo.RecordCount + oRsDebitMemo.RecordCount;

                if (DocsCounter > 0)
                {
                    #region Enviando Facturas de Venta

                    if (oRsInvoice.RecordCount > 0)
                    {
                        DllFunciones.Logger("Enviando facturas de venta...", sPathFileLog);

                        oRsInvoice.MoveFirst();

                        do
                        {
                            sDocEntry = null;
                            sDocEntry = Convert.ToString(oRsInvoice.Fields.Item("DocEntry").Value.ToString());
                            EnviarDocumentosMasivamenteTFHKA(oCompany, "FacturaDeClientes", "S", sDocEntry, sPathFileLog);

                            oRsInvoice.MoveNext();

                        } while (oRsInvoice.EoF == false);

                    }

                    #endregion

                    #region Enviando Notas Credito de Venta

                    if (oRsCreditMemo.RecordCount > 0)
                    {
                        DllFunciones.Logger("Enviando notas credito de clientes...", sPathFileLog);
                        oRsCreditMemo.MoveFirst();

                        do
                        {
                            sDocEntry = null;
                            sDocEntry = Convert.ToString(oRsCreditMemo.Fields.Item("DocEntry").Value.ToString());
                            EnviarDocumentosMasivamenteTFHKA(oCompany, "NotaCreditoClientes", "S", sDocEntry, sPathFileLog);
                            oRsCreditMemo.MoveNext();

                        } while (oRsCreditMemo.EoF == false);

                    }
                    #endregion

                    #region Enviando Notas Debito de Cliente

                    if (oRsDebitMemo.RecordCount > 0)
                    {
                        DllFunciones.Logger("Enviando notas debito de clientes...", sPathFileLog);
                        oRsDebitMemo.MoveFirst();

                        do
                        {
                            sDocEntry = null;
                            sDocEntry = Convert.ToString(oRsDebitMemo.Fields.Item("DocEntry").Value.ToString());
                            EnviarDocumentosMasivamenteTFHKA(oCompany, "NotaDebitoClientes", "S", sDocEntry, sPathFileLog);
                            oRsDebitMemo.MoveNext();

                        } while (oRsDebitMemo.EoF == false);

                    }

                    #endregion
                }

            }
            catch (Exception e)
            {
                DllFunciones.Logger(e.ToString(), sPathFileLog);
            }


        }

        public void InsertSendEmail(SAPbobsCOM.Company _oCompany, SAPbobsCOM.Recordset oDocs, BOeBillingService.ServicioEmisionFE.DocumentStatusResponse sRespuesta, string sPathFileLog, string _EstadoPDF)
        {
            Funciones.Comunes DllFunciones = new Funciones.Comunes();
            try
            {
                #region Variables y objetos

                string sCantidadCorreos = null;
                string sCorreo1 = null;
                string sCorreo2 = null;
                string sCorreo3 = null;
                string sCorreo4 = null;
                string sCorreo5 = null;
                int iContador = 0;

                if (sRespuesta == null)
                {
                    sCantidadCorreos = "0";
                }
                else
                {
                    sCantidadCorreos = Convert.ToString(sRespuesta.historialDeEntregas.Length);
                }

                iContador = Convert.ToInt32(oDocs.Fields.Item("Contador").Value.ToString());

                iContador++;

                #endregion

                #region Inserta el correo en la tablas de correos

                #region Variables y objetos

                string _sSerachNextCode = null;

                SAPbobsCOM.Recordset oSerachNextCode = (SAPbobsCOM.Recordset)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                _sSerachNextCode = DllFunciones.GetStringXMLDocument(_oCompany, "BOeBillingService", "eBilling", "SerachNextCode");

                oSerachNextCode.DoQuery(_sSerachNextCode);

                #endregion

                #region Asignacion de valores

                SAPbobsCOM.UserTable oUserTable;
                oUserTable = _oCompany.UserTables.Item("BOEE");

                if (string.IsNullOrEmpty(Convert.ToString(oDocs.Fields.Item("Code").Value.ToString())))
                {
                    oUserTable.Code = Convert.ToString(oSerachNextCode.Fields.Item("ID").Value.ToString());
                    oUserTable.Name = Convert.ToString(oSerachNextCode.Fields.Item("ID").Value.ToString());
                    oUserTable.UserFields.Fields.Item("U_BO_DocEntry").Value = Convert.ToString(oDocs.Fields.Item("DocEntry").Value.ToString());
                    oUserTable.UserFields.Fields.Item("U_BO_ObjecType").Value = Convert.ToString(oDocs.Fields.Item("ObjType").Value.ToString());

                }
                else
                {
                    oUserTable.GetByKey(Convert.ToString(oDocs.Fields.Item("Code").Value.ToString()));
                }

                if (sRespuesta == null)
                {

                }
                else
                {
                    if (sCantidadCorreos == "0")
                    {

                    }
                    else
                    {
                        oUserTable.UserFields.Fields.Item("U_BO_StatusEmail").Value = Convert.ToString(sRespuesta.historialDeEntregas[0].entregaEstatus.ToString());
                    }

                    oUserTable.UserFields.Fields.Item("U_BO_Count").Value = Convert.ToString(iContador);
                }

                if (_EstadoPDF == null)
                {

                }
                else
                {
                    oUserTable.UserFields.Fields.Item("U_BO_PdfTfhka").Value = _EstadoPDF;
                }


                #region Asignacion Correo 1

                if (sCantidadCorreos == "1")
                {
                    sCorreo1 = Convert.ToString(sRespuesta.historialDeEntregas[0].email.GetValue(0));

                    oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = sCorreo1;
                    oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = "";

                }

                #endregion

                #region Asignacion Correo 2

                if (sCantidadCorreos == "2")
                {
                    sCorreo1 = Convert.ToString(sRespuesta.historialDeEntregas[0].email.GetValue(0));
                    sCorreo2 = Convert.ToString(sRespuesta.historialDeEntregas[1].email.GetValue(0));

                    oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = sCorreo1;
                    oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = sCorreo2;
                    oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = "";


                }

                #endregion

                #region Asignacion Correo 3

                if (sCantidadCorreos == "3")
                {
                    sCorreo1 = Convert.ToString(sRespuesta.historialDeEntregas[0].email.GetValue(0));
                    sCorreo2 = Convert.ToString(sRespuesta.historialDeEntregas[1].email.GetValue(0));
                    sCorreo3 = Convert.ToString(sRespuesta.historialDeEntregas[2].email.GetValue(0));

                    oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = sCorreo1;
                    oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = sCorreo2;
                    oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = sCorreo3;
                    oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = "";
                    oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = "";


                }

                #endregion

                #region Asignacion Correo 4

                if (sCantidadCorreos == "4")
                {
                    sCorreo1 = Convert.ToString(sRespuesta.historialDeEntregas[0].email.GetValue(0));
                    sCorreo2 = Convert.ToString(sRespuesta.historialDeEntregas[1].email.GetValue(0));
                    sCorreo3 = Convert.ToString(sRespuesta.historialDeEntregas[2].email.GetValue(0));
                    sCorreo4 = Convert.ToString(sRespuesta.historialDeEntregas[3].email.GetValue(0));

                    oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = sCorreo1;
                    oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = sCorreo2;
                    oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = sCorreo3;
                    oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = sCorreo4;
                    oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = "";

                }

                #endregion

                #region Asignacion Correo 5

                if (sCantidadCorreos == "5")
                {
                    sCorreo1 = Convert.ToString(sRespuesta.historialDeEntregas[0].email.GetValue(0));
                    sCorreo2 = Convert.ToString(sRespuesta.historialDeEntregas[1].email.GetValue(0));
                    sCorreo3 = Convert.ToString(sRespuesta.historialDeEntregas[2].email.GetValue(0));
                    sCorreo4 = Convert.ToString(sRespuesta.historialDeEntregas[3].email.GetValue(0));
                    sCorreo5 = Convert.ToString(sRespuesta.historialDeEntregas[4].email.GetValue(0));

                    oUserTable.UserFields.Fields.Item("U_BO_Email1").Value = sCorreo1;
                    oUserTable.UserFields.Fields.Item("U_BO_Email2").Value = sCorreo2;
                    oUserTable.UserFields.Fields.Item("U_BO_Email3").Value = sCorreo3;
                    oUserTable.UserFields.Fields.Item("U_BO_Email4").Value = sCorreo4;
                    oUserTable.UserFields.Fields.Item("U_BO_Email5").Value = sCorreo5;

                }

                #endregion

                if (string.IsNullOrEmpty(Convert.ToString(oDocs.Fields.Item("Code").Value.ToString())))
                {
                    oUserTable.Add();
                }
                else
                {
                    oUserTable.Update();
                }

                #endregion

                #endregion

                #region Liberar Objetos

                DllFunciones.liberarObjetos(oSerachNextCode);
                DllFunciones.liberarObjetos(oUserTable);
                sRespuesta = null;

                #endregion

            }
            catch (Exception ex)
            {

                DllFunciones.Logger(ex.ToString(), sPathFileLog);
            }
        }

    }
}

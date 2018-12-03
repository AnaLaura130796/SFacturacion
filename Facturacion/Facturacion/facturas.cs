using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Facturacion
{
    class facturas
    {
        public static void tabla_facturacion()
        {
            //Dejamos listo en "sheetExportacion" nuestra hoja de excel para hacer los reportes.
            facturas.getPaths();
            string select_query = string.Format("SELECT " +
                " InvoiceCreditFlag " +
                /*", F2 as InvoiceDate " +
                ", F3 as InvoiceNo " +
                ", F4 as PoNo " +
                ", F5 as Currency " +
                ", '' as BIL " +
                ", F6 as SapBox " +
                ", '' as BaselineDate " +
                ", '' as ShipToCountry " +
                ", '' as ShipFromCountry " +
                ", F7 as LegalEntity " +
                ", F8 as LE_Country" +
                ", F9 as Name_Lab " +
                ", '' as PGVAT_ID " + 
                ", F10 as Address " +
                ", F11 as Address1 " +
                ", F12 as City1 " +
                ", F13 as PostalCode " +
                ", F14 as Country " +
                ", F15 as Vendor_No " +
                ", F16 Name_Vendor " +
                ", '' as PartnerVAT " +
                ", F17 as Address1" +
                ", F18 as Address2 " +
                ", F19 as City1 " +
                ", F20 as Country " +
                ", '' as BankCountryKey " +
                ", '' as BankAccountNo " +
                ", '' as RemitTo_No " +
                ", '' as Name " +
                ", '' as Address " +
                ", '' as Address1 " +
                ", '' as City1 " +
                ", '' as PostalCode " +
                ", '' as Country " +
                ", '' as ProfitCenter " +
                ", '' as WbsElement " +
                ", '' as customeServiceOrder " +
                ", '' as Gobal_Bussinerss_Area " +
                ", '' as PaymentMethod " +
                ", '' as PaymentMethod_Supplement " +
                ", '' as item_Text " +
                ", '' as PartDescription " +
                ", F21 as LineItemNumber " +
                ", F22 as Quantity " +
                ", F23 as UoM " +
                ", F24 as NetPrice " +
                ", F25 as LineItemAmount " +
                ", '' as PoNo " +
                ", F26 as MaterialNumber " +
                ", '' as TaxID " +
                ", '' as TaxAmount " +
                ", '' as TaxRate " +
                ", '' as UnplannedDeliveryCost " +
                ", '' as ISR_No " +
                ", '' as ISR_Ref_No " +
                ", '' as ScbIndicator " +
                ", '' as WithHoldingTax " +
                ", '' as NetAmount " +
                ", F27 as InvoiceAmount " +
                ", '' as TaxID " +
                ", '' as TaxType " +
                ", F28 as TaxAmount " +
                ", F29 as TaxRate " +
                ", '' as AllowanceAmount " +
                ", '' as AllowanceDescription " +
                ", '' as AllowanceCode " +
                ", '' as ChargesAmount " +
                ", '' as ChargesDescription " +
                ", '' as ChargesCode " +
                ", F30 as InvoiceType " +
                ", '' as ShipToName " +
                ", '' as ShipToAddress " +
                ", '' as ShipToState " +
                ", '' as ShipToCity1 " +
                ", '' as ShipToPostalCode " +
                ", '' as ShipToCountry " +
                ", '' as TaxId1 " +
                ", '' as TaxType1 " +
                ", '' as TaxRate1 " +
                ", '' as TaxAmount1 " +*/
                     "FROM  [{0}$]", "hoja_entrada");
            DataTable tabla_facturacion = DataBase.runSelectQuery(select_query);

            if (tabla_facturacion == null)
                return; 

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(facturas.pathSolicitudFacturaciones);
            Microsoft.Office.Interop.Excel.Workbook workbookExportacion;
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkBook.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            xlWorkSheet.Copy(Type.Missing, Type.Missing);
            xlWorkBook.Saved = true;
            xlWorkBook.Close(0);
            workbookExportacion = xlApp.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Sheets sheetsExportacion = workbookExportacion.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet sheetExportacion = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            //se va a guardar la información existente en la plantilla de entrada
           

            //Copiamos la información dentro de la solicitud. 
            sheetExportacion.Cells[1, 1] = "HOLA"; 
          sheetExportacion.Cells[2, 1] = tabla_facturacion.Rows[0]["InvoiceCreditFlag"].ToString();
          sheetExportacion.Cells[2, 2] = tabla_facturacion.Rows[0]["InvoiceDate"].ToString();
          sheetExportacion.Cells[2, 3] = tabla_facturacion.Rows[0]["InvoiceNo"].ToString();
          sheetExportacion.Cells[2, 4] = tabla_facturacion.Rows[0]["PoNo"].ToString();
          sheetExportacion.Cells[2, 5] = tabla_facturacion.Rows[0]["Currency"].ToString();
          sheetExportacion.Cells[2, 6] = tabla_facturacion.Rows[0]["SapBox"].ToString();
          sheetExportacion.Cells[2, 7] = tabla_facturacion.Rows[0]["LegalEntity"].ToString();
          sheetExportacion.Cells[2, 8] = tabla_facturacion.Rows[0]["LE_Country"].ToString();
          sheetExportacion.Cells[2, 9] = tabla_facturacion.Rows[0]["Name_Lab"].ToString();
          sheetExportacion.Cells[2, 10] = tabla_facturacion.Rows[0]["Address"].ToString();
          sheetExportacion.Cells[2, 11] = tabla_facturacion.Rows[0]["Address1"].ToString();
          sheetExportacion.Cells[2, 12] = tabla_facturacion.Rows[0]["City1"].ToString();
          sheetExportacion.Cells[2, 13] = tabla_facturacion.Rows[0]["PostalCode"].ToString();
          sheetExportacion.Cells[2, 14] = tabla_facturacion.Rows[0]["Country"].ToString();
          sheetExportacion.Cells[2, 15] = tabla_facturacion.Rows[0]["Vendor_No"].ToString();
          sheetExportacion.Cells[2, 16] = tabla_facturacion.Rows[0]["Name_Vendor"].ToString();
          sheetExportacion.Cells[2, 17] = tabla_facturacion.Rows[0]["Address_vendor"].ToString();
          sheetExportacion.Cells[2, 18] = tabla_facturacion.Rows[0]["Address2_vendor"].ToString();
          sheetExportacion.Cells[2, 19] = tabla_facturacion.Rows[0]["City1_vendor"].ToString();
          sheetExportacion.Cells[2, 20] = tabla_facturacion.Rows[0]["Country_vendor"].ToString();
          sheetExportacion.Cells[2, 21] = tabla_facturacion.Rows[0]["LineItemNumber"].ToString();
          sheetExportacion.Cells[2, 22] = tabla_facturacion.Rows[0]["Quantity"].ToString();
          sheetExportacion.Cells[2, 23] = tabla_facturacion.Rows[0]["UoM"].ToString();
          sheetExportacion.Cells[2, 24] = tabla_facturacion.Rows[0]["NetPrice"].ToString();
          sheetExportacion.Cells[2, 25] = tabla_facturacion.Rows[0]["LineItemAmount"].ToString();
          sheetExportacion.Cells[2, 26] = tabla_facturacion.Rows[0]["MaterialNumber"].ToString();
          sheetExportacion.Cells[2, 27] = tabla_facturacion.Rows[0]["InvoiceAmount"].ToString();
          sheetExportacion.Cells[2, 28] = tabla_facturacion.Rows[0]["TaxAmount"].ToString();
          sheetExportacion.Cells[2, 29] = tabla_facturacion.Rows[0]["TaxRate"].ToString();
          sheetExportacion.Cells[2, 30] = tabla_facturacion.Rows[0]["InvoiceType"].ToString();


            //Establecemos los rangos de impresión.             
           sheetExportacion.PageSetup.PrintArea = "A1:AD30";
            
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;
          /*  string rutaPDF = System.Windows.Forms.Application.StartupPath + "\\ultimoReporte.pdf";
            //MessageBox.Show("Guardado en " + rutaPDF); 
            sheetExportacion.ExportAsFixedFormat(
            Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
            rutaPDF,
            Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
            true,
            false,
            Type.Missing,
            Type.Missing,
            false);
            xlApp.ActiveWorkbook.Saved = true;
            xlApp.ActiveWorkbook.Close(0);
            utilidades.matarProcesoDeExcel(xlApp);
        //    utilidades.verPDF(rutaPDF); */
        }

        private static void getPaths()
        {
            string[] lines = new string[0];
            try
            {

                lines = System.IO.File.ReadAllLines(@"" + System.Windows.Forms.Application.StartupPath + "\\paths.txt");
                pathSolicitudFacturaciones = lines[0];
            }
            catch
            {
                utilidades.mostrarMensajeValidacion("No se pudo leer la configuración. Verificar el archivo de paths.");
                System.Environment.Exit(1);
            }
        }

        public static string pathSolicitudFacturaciones { get; set; }

     
    }
}

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
        public static void solicitud_facturacion(long solicitud)
        {
            //Dejamos listo en "sheetExportacion" nuestra hoja de excel para hacer los reportes.
            facturas.getPaths();
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
            string select_query = string.Format("select " +
               "InvoiceCreditFlag " +
                ", InvoiceDate " +
                ", InvoiceNo " +
                ", PoNo " +
                ", Currency " +
                ", SapBox " +
                ", LegalEntity " +
                ", LE_Country" +
                ", Name " +
                ", Address1 " +
                ", Address2 " +
                ", City1 " +
                ", Country " +
                ", LineItemNumber " +
                ", Quantity " +
                ", UoM " +
                ", NetPrice " +
                ", LineItemAmount " +
                ", MaterialNumber " +
                ", InvoiceAmount " +
                ", TaxAmount " +
                ", TaxRate " +
                ", InvoiceType " +
                "FROM  [{0}$] ", "hoja_entrada");
            DataTable solicitud_facturacion = DataBase.runSelectQuery(select_query);
            //Copiamos la información dentro de la solicitud. 
          sheetExportacion.Cells[1, 1] = solicitud;
            sheetExportacion.Cells[2, 1] = solicitud_facturacion.Rows[0]["InvoiceCreditFlag"].ToString();
            sheetExportacion.Cells[2, 2] = solicitud_facturacion.Rows[0]["InvoiceDate"].ToString();
            sheetExportacion.Cells[2, 3] = solicitud_facturacion.Rows[0]["InvoiceNo"].ToString();
            sheetExportacion.Cells[2, 4] = solicitud_facturacion.Rows[0]["PoNo"].ToString();
            sheetExportacion.Cells[2, 5] = solicitud_facturacion.Rows[0]["SapBox"].ToString();
            sheetExportacion.Cells[2, 6] = solicitud_facturacion.Rows[0]["Currency"].ToString();
            sheetExportacion.Cells[2, 7] = solicitud_facturacion.Rows[0]["LegalEntity"].ToString();
            sheetExportacion.Cells[2, 8] = solicitud_facturacion.Rows[0]["LE_Country"].ToString();
            sheetExportacion.Cells[2, 9] = solicitud_facturacion.Rows[0]["Name"].ToString();
            sheetExportacion.Cells[2, 10] = solicitud_facturacion.Rows[0]["Address1"].ToString();
            sheetExportacion.Cells[2, 11] = solicitud_facturacion.Rows[0]["Address2"].ToString();
            sheetExportacion.Cells[2, 12] = solicitud_facturacion.Rows[0]["City1"].ToString();
            sheetExportacion.Cells[2, 13] = solicitud_facturacion.Rows[0]["Country"].ToString();
            sheetExportacion.Cells[2, 14] = solicitud_facturacion.Rows[0]["LineItemNumber"].ToString();
            sheetExportacion.Cells[2, 15] = solicitud_facturacion.Rows[0]["Quantity"].ToString();
            sheetExportacion.Cells[2, 16] = solicitud_facturacion.Rows[0]["UoM"].ToString();
            sheetExportacion.Cells[2, 17] = solicitud_facturacion.Rows[0]["NetPrice"].ToString();
            sheetExportacion.Cells[2, 18] = solicitud_facturacion.Rows[0]["LineItemAmount"].ToString();
            sheetExportacion.Cells[2, 19] = solicitud_facturacion.Rows[0]["MaterialNumber"].ToString();
            sheetExportacion.Cells[2, 20] = solicitud_facturacion.Rows[0]["InvoiceAmount"].ToString();
            sheetExportacion.Cells[2, 21] = solicitud_facturacion.Rows[0]["TaxAmount"].ToString();
            sheetExportacion.Cells[2, 22] = solicitud_facturacion.Rows[0]["TaxRate"].ToString();
            sheetExportacion.Cells[2, 23] = solicitud_facturacion.Rows[0]["InvoiceType"].ToString();

            //Establecemos los rangos de impresión.             
           sheetExportacion.PageSetup.PrintArea = "A1:AD30";
            
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;
            string rutaPDF = System.Windows.Forms.Application.StartupPath + "\\ultimoReporte.pdf";
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
        //    utilidades.verPDF(rutaPDF); 
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

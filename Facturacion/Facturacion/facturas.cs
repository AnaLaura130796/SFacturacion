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
                "F1 as InvoiceCreditFlag " +
                ", F2 as InvoiceDate " +
                ", F3 as InvoiceNo " +
                ", F4 as PoNo_1 " +
                ", F5 as Currency_1 " +
                ", '' as BIL " +
                ", F6 as SapBox " +
                ", '' as BaselineDate " +
                ", '' as ShipToCountry_A " +
                ", '' as ShipFromCountry " +
                ", F7 as LegalEntity_1 " +
                ", F8 as LE_Country" +
                ", F9 as Name_Lab " +
                ", '' as PGVAT_ID " + 
                ", F10 as Address_a " +
                ", F11 as Address1_1 " +
                ", F12 as City1_1 " +
                ", F13 as PostalCode_A " +
                ", F14 as Country_1 " +
                ", F15 as Vendor_No " +
                ", F16 as Name_Vendor " +
                ", ''  as PartnerVAT " +
                ", F17 as Address1_2" +
                ", F18 as Address2 " +
                ", F19 as City1_2 " +
                ", F20 as Country_2 " +
                ", '' as BankCountryKey " +
                ", '' as BankAccountNo " +
                ", '' as RemitTo_No " +
                ", '' as Name " +
                ", '' as Address_b " +
                ", '' as Address1_3 " +
                ", '' as City1_3 " +
                ", '' as PostalCode_B " +
                ", '' as Country_3 " +
                ", '' as ProfitCenter " +
                ", '' as WbsElement " +
                ", '' as customeServiceOrder " +
                ", '' as Gobal_Bussinerss_Area " +
                ", '' as PaymentMethod " +
                ", '' as PaymentMethod_Supplement " +
                ", '' as item_Text " +
                ", '' as PartDescription " +
                ", F21 as LineItemNumber " +
                ", F22 as Quantity_1 " +
                ", F23 as UoM " +
                ", F24 as NetPrice " +
                ", F25 as LineItemAmount " +
                ", '' as PoNo_2 " +
                ", F26 as MaterialNumber " +
                ", '' as TaxID_A " +
                ", '' as TaxAmount_1 " +
                ", '' as TaxRate_a " +
                ", '' as UnplannedDeliveryCost " +
                ", '' as ISR_No " +
                ", '' as ISR_Ref_No " +
                ", '' as ScbIndicator " +
                ", '' as WithHoldingTax " +
                ", '' as NetAmount " +
                ", F27 as InvoiceAmount " +
                ", '' as TaxID_B " +
                ", '' as TaxType " +
                ", F28 as TaxAmount_2 " +
                ", F29 as TaxRate_b " +
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
                ", '' as ShipToCountry_B " +
                ", '' as TaxId1 " +
                ", '' as TaxType1 " +
                ", '' as TaxRate1 " +
                ", '' as TaxAmount1 " +
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
            sheetExportacion.Cells[2, 4] = tabla_facturacion.Rows[0]["PoNo_1"].ToString(); 
            sheetExportacion.Cells[2, 5] = tabla_facturacion.Rows[0]["Currency_1"].ToString(); 
            sheetExportacion.Cells[2, 6] = tabla_facturacion.Rows[0]["BIL"].ToString(); 
            sheetExportacion.Cells[2, 7] = tabla_facturacion.Rows[0]["SapBox"].ToString();
            sheetExportacion.Cells[2, 8] = tabla_facturacion.Rows[0]["BaselineDate"].ToString();
            sheetExportacion.Cells[2, 9] = tabla_facturacion.Rows[0]["ShipToCountry_A"].ToString();
            sheetExportacion.Cells[2, 10] = tabla_facturacion.Rows[0]["ShipFromCountry"].ToString();
            sheetExportacion.Cells[2, 11] = tabla_facturacion.Rows[0]["LegalEntity_1"].ToString();
            sheetExportacion.Cells[2, 12] = tabla_facturacion.Rows[0]["LE_Country"].ToString();
            sheetExportacion.Cells[2, 13] = tabla_facturacion.Rows[0]["Name_Lab"].ToString();
            sheetExportacion.Cells[2, 14] = tabla_facturacion.Rows[0]["PGVAT_ID"].ToString();
            sheetExportacion.Cells[2, 15] = tabla_facturacion.Rows[0]["Address_a"].ToString(); 
            sheetExportacion.Cells[2, 16] = tabla_facturacion.Rows[0]["Address1_1"].ToString();
            sheetExportacion.Cells[2, 17] = tabla_facturacion.Rows[0]["City1_1"].ToString();
            sheetExportacion.Cells[2, 18] = tabla_facturacion.Rows[0]["PostalCode_A"].ToString();
            sheetExportacion.Cells[2, 19] = tabla_facturacion.Rows[0]["Country_1"].ToString();
            sheetExportacion.Cells[2,20] = tabla_facturacion.Rows[0]["Vendor_No"].ToString();
            sheetExportacion.Cells[2, 21] = tabla_facturacion.Rows[0]["Name_Vendor"].ToString();
            sheetExportacion.Cells[2, 22] = tabla_facturacion.Rows[0]["PartnerVAT"].ToString();
            sheetExportacion.Cells[2, 23] = tabla_facturacion.Rows[0]["Address1_2"].ToString();
            sheetExportacion.Cells[2, 24] = tabla_facturacion.Rows[0]["Address2"].ToString();
            sheetExportacion.Cells[2, 25] = tabla_facturacion.Rows[0]["City1_2"].ToString();
            sheetExportacion.Cells[2, 26] = tabla_facturacion.Rows[0]["Country_2"].ToString();
            sheetExportacion.Cells[2, 27] = tabla_facturacion.Rows[0]["BankCountryKey"].ToString();
            sheetExportacion.Cells[2, 28] = tabla_facturacion.Rows[0]["BankAccountNo"].ToString();
            sheetExportacion.Cells[2, 29] = tabla_facturacion.Rows[0]["RemitTo_No"].ToString();
            sheetExportacion.Cells[2, 30] = tabla_facturacion.Rows[0]["Name"].ToString();
            sheetExportacion.Cells[2, 31] = tabla_facturacion.Rows[0]["Address_b"].ToString();
            sheetExportacion.Cells[2, 32] = tabla_facturacion.Rows[0]["Address1_3"].ToString();
            sheetExportacion.Cells[2, 33] = tabla_facturacion.Rows[0]["City1_3"].ToString();
            sheetExportacion.Cells[2, 34] = tabla_facturacion.Rows[0]["PostalCode_B"].ToString();
            sheetExportacion.Cells[2, 35] = tabla_facturacion.Rows[0]["Country_3"].ToString();
            sheetExportacion.Cells[2, 36] = tabla_facturacion.Rows[0]["ProfitCenter"].ToString();
            sheetExportacion.Cells[2, 37] = tabla_facturacion.Rows[0]["WbsElement"].ToString();
            sheetExportacion.Cells[2, 38] = tabla_facturacion.Rows[0]["customeServiceOrder"].ToString();
            sheetExportacion.Cells[2, 39] = tabla_facturacion.Rows[0]["Gobal_Bussinerss_Area"].ToString();
            sheetExportacion.Cells[2, 40] = tabla_facturacion.Rows[0]["PaymentMethod"].ToString();
            sheetExportacion.Cells[2, 41] = tabla_facturacion.Rows[0]["PaymentMethod_Supplement"].ToString();
            sheetExportacion.Cells[2, 42] = tabla_facturacion.Rows[0]["item_Text"].ToString();
            sheetExportacion.Cells[2, 43] = tabla_facturacion.Rows[0]["PartDescription"].ToString();
            sheetExportacion.Cells[2, 44] = tabla_facturacion.Rows[0]["LineItemNumber"].ToString();
            sheetExportacion.Cells[2, 45] = tabla_facturacion.Rows[0]["Quantity_1"].ToString();
            sheetExportacion.Cells[2, 46] = tabla_facturacion.Rows[0]["UoM"].ToString();
            sheetExportacion.Cells[2, 47] = tabla_facturacion.Rows[0]["NetPrice"].ToString();
            sheetExportacion.Cells[2, 48] = tabla_facturacion.Rows[0]["LineItemAmount"].ToString();
            sheetExportacion.Cells[2, 49] = tabla_facturacion.Rows[0]["PoNo_2"].ToString();
            sheetExportacion.Cells[2, 50] = tabla_facturacion.Rows[0]["MaterialNumber"].ToString();
            sheetExportacion.Cells[2, 51] = tabla_facturacion.Rows[0]["TaxID_A"].ToString();
            sheetExportacion.Cells[2, 52] = tabla_facturacion.Rows[0]["TaxAmount_1"].ToString();
            sheetExportacion.Cells[2, 53] = tabla_facturacion.Rows[0]["TaxRate_a"].ToString();
            sheetExportacion.Cells[2, 54] = tabla_facturacion.Rows[0]["UnplannedDeliveryCost"].ToString();
            sheetExportacion.Cells[2, 55] = tabla_facturacion.Rows[0]["ISR_No"].ToString();
            sheetExportacion.Cells[2, 56] = tabla_facturacion.Rows[0]["ISR_Ref_No"].ToString();
            sheetExportacion.Cells[2, 57] = tabla_facturacion.Rows[0]["ScbIndicator"].ToString();
            sheetExportacion.Cells[2, 58] = tabla_facturacion.Rows[0]["WithHoldingTax"].ToString();
            sheetExportacion.Cells[2, 59] = tabla_facturacion.Rows[0]["NetAmount"].ToString();
            sheetExportacion.Cells[2, 60] = tabla_facturacion.Rows[0]["InvoiceAmount"].ToString();
            sheetExportacion.Cells[2, 61] = tabla_facturacion.Rows[0]["TaxID_B"].ToString();
            sheetExportacion.Cells[2, 62] = tabla_facturacion.Rows[0]["TaxType"].ToString();
            sheetExportacion.Cells[2, 63] = tabla_facturacion.Rows[0]["TaxAmount_2"].ToString();
            sheetExportacion.Cells[2, 64] = tabla_facturacion.Rows[0]["TaxRate_b"].ToString();
            sheetExportacion.Cells[2, 65] = tabla_facturacion.Rows[0]["AllowanceAmount"].ToString();
            sheetExportacion.Cells[2, 66] = tabla_facturacion.Rows[0]["AllowanceDescription"].ToString();
            sheetExportacion.Cells[2, 67] = tabla_facturacion.Rows[0]["AllowanceCode"].ToString();
            sheetExportacion.Cells[2, 68] = tabla_facturacion.Rows[0]["ChargesAmount"].ToString();
            sheetExportacion.Cells[2, 69] = tabla_facturacion.Rows[0]["ChargesDescription"].ToString();
            sheetExportacion.Cells[2, 70] = tabla_facturacion.Rows[0]["ChargesCode"].ToString();
            sheetExportacion.Cells[2, 71] = tabla_facturacion.Rows[0]["InvoiceType"].ToString();
            sheetExportacion.Cells[2, 72] = tabla_facturacion.Rows[0]["ShipToName"].ToString();
            sheetExportacion.Cells[2, 73] = tabla_facturacion.Rows[0]["ShipToAddress"].ToString();
            sheetExportacion.Cells[2, 74] = tabla_facturacion.Rows[0]["ShipToState"].ToString();
            sheetExportacion.Cells[2, 75] = tabla_facturacion.Rows[0]["ShipToCity1"].ToString();
            sheetExportacion.Cells[2, 76] = tabla_facturacion.Rows[0]["ShipToPostalCode"].ToString();
            sheetExportacion.Cells[2, 77] = tabla_facturacion.Rows[0]["ShipToCountry_B"].ToString();
            sheetExportacion.Cells[2, 78] = tabla_facturacion.Rows[0]["TaxId1"].ToString();
            sheetExportacion.Cells[2, 79] = tabla_facturacion.Rows[0]["TaxType1"].ToString();
            sheetExportacion.Cells[2, 80] = tabla_facturacion.Rows[0]["TaxRate1"].ToString();
            sheetExportacion.Cells[2, 81] = tabla_facturacion.Rows[0]["TaxAmount1"].ToString();
            MessageBox.Show("ya lei los datos");
            //Establecemos los rangos de impresión.             
           sheetExportacion.PageSetup.PrintArea = "A1:AD30";
            xlApp.DisplayAlerts = false;
            xlApp.Visible = false;
            string rutaPDF = System.Windows.Forms.Application.StartupPath + "\\ultimoReporte.pdf";
            MessageBox.Show("Guardado en " + rutaPDF); 
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
            MessageBox.Show("ya sali de facturas.tabla_facturacion() el valor de tabla_facturacion es" + select_query);
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
  

        public static void generarFactura(string encabezado, DataTable tabla)
        {
            //Creamos un objeto misvalue para facilitar las configuraciones. 
            object misValue = System.Reflection.Missing.Value;

            //Creamos una nueva aplicación de excel. 
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();


            //Abrimos la plantilla de reportes y creamos un nuevo workbook para mostrar ahí el reporte. 

            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(facturas.pathPlantillaMantenimientosSemanales);
            Microsoft.Office.Interop.Excel.Workbook workbookExportacion;
            //Obtenemos todas las hojas de la plantilla 
            Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkBook.Worksheets;

            //Obtenemos la primera hoja de la plantilla 
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

            //Se copia la hoja actual y se coloca automáticamente en un nuevo workbook. 
            xlWorkSheet.Copy(Type.Missing, Type.Missing);

            //Cerramos la plantilla original de excel. 
            xlWorkBook.Close(0);


            //Asignamos un identificador para nuestro workbook recien creado. 
            workbookExportacion = xlApp.ActiveWorkbook;
            Microsoft.Office.Interop.Excel.Sheets sheetsExportacion = workbookExportacion.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet sheetExportacion = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;



            //Pegamos la tabla de la exportación en la celda 8, 1
            Microsoft.Office.Interop.Excel.Range CR = sheetExportacion.Cells[8, 1] as Microsoft.Office.Interop.Excel.Range;
            CR.Select();
            sheetExportacion.Paste(CR, Clipboard.GetText());
        }


        public static string pathPlantillaMantenimientosSemanales { get; set; }
    }
}

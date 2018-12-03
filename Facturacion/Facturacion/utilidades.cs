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

    class utilidades
    {
        
        public static void mostrarMensajeValidacion(string mensaje)
        {
            MessageBox.Show(mensaje,
            "Aviso",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error,
            MessageBoxDefaultButton.Button1);
        }
        internal static void exportarTablaExcel(DataTable tabla, string encabezado)
        {

            try
            {
              
                    if (tabla == null)
                    {
                        utilidades.mostrarMensajeValidacion("No se encontró información en la tabla para exportación. Contacta a Aseguramiento de calidad.");
                    }
            
                    //Creamos una nueva aplicación de excel. 
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                    //Abrimos la plantilla de reportes y creamos un nuevo workbook para mostrar ahí el reporte.
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
                    //Obtenemos todas las hojas de la plantilla 
                    Microsoft.Office.Interop.Excel.Sheets sheets = xlWorkBook.Worksheets;

                    //Obtenemos la primera hoja de la plantilla 
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = xlApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


                    //Copiamos la tabla en el portapapeles con el encabezado.
                    utilidades.CopyDataTableToClipboard(tabla, true);

                    //Pegamos nuestra tabla para la generación del reporte. 
                    Microsoft.Office.Interop.Excel.Range CR = xlWorkSheet.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                    CR.Select();
                    xlWorkSheet.Paste();

                    //Colocamos los bordes de las celdas 
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[tabla.Rows.Count + 1, tabla.Columns.Count]].borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[tabla.Rows.Count + 1, tabla.Columns.Count]].borders.Weight = 2d;

                    //Coloreamos los encabezados de las celdas. 
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, tabla.Columns.Count]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                    xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, tabla.Columns.Count]].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                    //Establecemos los márgenes para la impresión y hacemos autoFit
                    /*tabla.PageSetup.PrintArea = "A1:E" + ultimaCelda;
                    Microsoft.Office.Interop.Excel.Range aRange = sheetExportacion.get_Range("A7", "E" + ultimaCelda);
                    aRange.Rows.AutoFit();
                     * */
                 /*   string rutaPDF = System.Windows.Forms.Application.StartupPath + "\\ultimoReporte.pdf";
                    //MessageBox.Show("Guardado en " + rutaPDF); 
                    xlWorkSheet.ExportAsFixedFormat(
                    Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                    rutaPDF,
                    Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                    true,
                    false,
                    Type.Missing,
                    Type.Missing,
                    false);
                    xlApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMaximized;
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = true;
                    xlWorkBook.WindowDeactivate += cerrarExcel;*/
                
            }
            catch (Exception e)
            {
                utilidades.mostrarMensajeValidacion(e.Message.ToString());

            }
        }

        
        private static void cerrarExcel(Microsoft.Office.Interop.Excel.Window Wn)
        {
            Wn.Application.Quit();
            utilidades.matarProcesoDeExcel(Wn.Application);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static void matarProcesoDeExcel(Microsoft.Office.Interop.Excel.Application xlApp)
        {
            uint processId = 0;
            GetWindowThreadProcessId(new IntPtr(xlApp.Hwnd), out processId);
            try
            {
                if (processId != 0)
                {
                    Process excelProcess = Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
            catch
            {
                // Process was already killed
            }
        }

        internal static void CopyDataTableToClipboard(DataTable DT, bool headerCopied = false)
        {
            if (DT == null)
            {
                Clipboard.SetText(" ");
                return;
            }
            if ((DT.Rows.Count == 0))
            {
                Clipboard.SetText(" ");
                return;
            }
            StringBuilder Output = new StringBuilder();

            //The first "line" will be the Headers.
            if (headerCopied)
            {
                for (int i = 0; i < DT.Columns.Count; i++)
                {
                    Output.Append(DT.Columns[i].ColumnName + "\t");
                }
            }

            Output.Append("\n");

            //Generate Cell Value Data
            foreach (DataRow Row in DT.Rows)
            {
                for (int i = 0; i < Row.ItemArray.Length; i++)
                {
                    //Handling the last cell of the line.
                    if (i == (Row.ItemArray.Length - 1))
                    {

                        Output.Append(Row.ItemArray[i].ToString() + "\n");
                    }
                    else
                    {

                        Output.Append(Row.ItemArray[i].ToString() + "\t");
                    }
                }
            }

            Clipboard.SetText(Output.ToString());
        }


        
    }
}


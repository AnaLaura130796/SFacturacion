using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Facturacion
{
    public partial class FormFacturas : Form
    {


        public FormFacturas()
        {
            
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonSeleccion_Archivo_Click(object sender, EventArgs e)
        {
            //este botón utiliza OpenFileDialog para seleccionar el archivo de entrada
            var fileContent = string.Empty;
            var filePath = string.Empty;
            using (OpenFileDialog openfiledialog = new OpenFileDialog())
            {   //abrirá en la ruta donde esta el archivo a entrar
                openfiledialog.InitialDirectory = "C:\\Users\\PMM\\Desktop\\sistemas\\facturacion_excel";
                openfiledialog.RestoreDirectory = true;
                //solo muestra los archivos de este tipo
                openfiledialog.Filter = "xlsx files (*,.xlsx)|*.xlsx";
                if (openfiledialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openfiledialog.FileName;
                    DataBase.pathBaseDeDatos = filePath;
                }
            }

        }
        

        private void buttonLeer_Contenido_Click(object sender, EventArgs e)
        {

            leer_tabla();
        }

        private void buttonGenera_Excel_Click(object sender, EventArgs e)
        {
            leer_tabla();
            if (tabla_generada != null)
            {
                utilidades.exportarTablaExcel(tabla_facturacion);
                facturas.generarFactura("Facturas", tabla_generada);
            }
        }
        DataTable tabla_generada = new DataTable();
        DataTable tabla_facturacion;
        public void leer_tabla( )

        {
            if (DataBase.pathBaseDeDatos == null)
            {
                return ;
            }
            else
            {
                facturas.tabla_facturacion();
               // DataTable tabla_generada = new DataTable();
                string query_select = string.Format("SELECT " +
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
         tabla_facturacion= DataBase.runSelectQuery(query_select);
            
            }
         
        }








      

      
    }
}

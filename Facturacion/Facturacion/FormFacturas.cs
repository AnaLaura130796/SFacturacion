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
            using(OpenFileDialog openfiledialog = new OpenFileDialog())
            {   //abrirá en la ruta donde esta el archivo a entrar
                openfiledialog.InitialDirectory = "C:\\Users\\PMM\\Desktop\\sistemas\\facturacion_excel";
                openfiledialog.RestoreDirectory = true;
                //solo muestra los archivos de este tipo
                openfiledialog.Filter = "xlsx files (*,.xlsx)|*.xlsx";
                if(openfiledialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openfiledialog.FileName;
                    var filestream = openfiledialog.OpenFile();

                    using(StreamReader reader = new StreamReader(filestream))
                    {
                        fileContent = reader.ReadToEnd();
                        MessageBox.Show(fileContent, "Ruta del archivo seleccionado: " + filePath);
                    }

                }
            } 

        }

        private void buttonLeer_Contenido_Click(object sender, EventArgs e)
        {
            /*este boton ejecutará la función que guardará cada uno de los datos del archivo de entrada
             * ´guardándolos en una variable DataTable que será ejecutada por el DataBase para correr el query
             * el cuál será empleado después para generar el archivo de salida con un nuevo formato
             * agregando las columnas vacias que no están en el archivo de entrada.
            */
            //esta función copia los datos de la plantilla de entrada.
            facturas.solicitud_facturacion(this.solicitud);
            utilidades.exportarTablaExcel(tabla);          
        }

        private void buttonGenera_Excel_Click(object sender, EventArgs e)
        {
            /*Función que exporta los datos del portapapeles al nuevo excel */
            
            
              
            
            
           
        }




        public long solicitud { get; set; }

        public DataTable tabla { get; set; }
    }
}

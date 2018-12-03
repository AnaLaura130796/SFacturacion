using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Facturacion
{
    class DataBase
    {
        static public OleDbConnection _conexionFacturacion = null;

        public static void getStringConnectionFacturaciones()
        {            
            stringConnectionFacturaciones = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathBaseDeDatos + "; Extended Properties='Excel 8.0;HDR=NO';"; 
        }

        public static string stringConnectionFacturaciones { get; set; }

        internal static void crearConexion()
        {
            //Creamos la conexión 
            DataBase.getStringConnectionFacturaciones();
            DataBase._conexionFacturacion = new OleDbConnection(DataBase.stringConnectionFacturaciones);

        }

        internal static System.Data.DataTable runSelectQuery(string query)
        {
            try
            {

                //En caso de que la conexión no este inicializada, la creamos. 
                if (DataBase._conexionFacturacion == null)
                {
                    DataBase.crearConexion();
                }

                //Creamos el comando 
                OleDbCommand comando = new OleDbCommand();
                comando.Connection = DataBase._conexionFacturacion;

                //Creamos el DataAdapter
                OleDbDataAdapter adaptadorDeDatos = new OleDbDataAdapter();
                adaptadorDeDatos.SelectCommand = comando;


                DataSet ds = new DataSet();
                comando.CommandText = query;

                _conexionFacturacion.Open();
                DataTable tabla = new DataTable();
                adaptadorDeDatos.Fill(tabla);
                _conexionFacturacion.Close();

                if (tabla.Rows.Count == 0)
                {
                    return null;
                }
                return tabla;


            }
            catch (Exception ex)
            {
                ManageException(ex);
                return null;
            }
        }

        private static void ManageException(Exception ex)
        {
            if (_conexionFacturacion.State == System.Data.ConnectionState.Open)
                _conexionFacturacion.Close();
            MessageBox.Show(ex.ToString());
            MessageBox.Show("Es posible que la base de datos de Especificaciones no este disponible");
        }

        public static string pathBaseDeDatos { get; set; }
        
    }
}
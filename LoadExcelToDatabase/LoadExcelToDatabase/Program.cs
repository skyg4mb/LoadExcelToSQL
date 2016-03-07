using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace LoadExcelToDatabase
{
    class Program
    {
        static void Main(string[] args)
        {
            //Prodimiento que carga la informacion de chat web
            CargueABaseDeDatos(@"Directorio", "Tabla", "Tabla Temporal", "TablaHIstorica", "postejecucion", "nombre proceso");
           
        }

        private static void CargueABaseDeDatos(String directorio, string FileName, string tablaDestino, string tablaControl, string procedimientoFinal, string HojaExcel)

        {

            DirectoryInfo directory = new DirectoryInfo(@directorio);
            FileInfo[] files = directory.GetFiles(FileName+"*");
            DataTable dt;

            try{
                for (int i = 0; i < files.Length; i++)
                {
                    Console.WriteLine(((FileInfo)files[i]).FullName);
                    dt=CargarExcelDatatable(((FileInfo)files[i]).FullName, HojaExcel);
                    CargueABaseDeDatos(dt, ((FileInfo)files[i]).FullName, tablaDestino, tablaControl, procedimientoFinal);
                }
            }
            catch (Exception ex)

            {

            Console.WriteLine(ex.ToString());

            }

            

        }

        private static DataTable CargarExcelDatatable(string ruta, string HojaExcel)
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ruta + 
                                        @";Extended Properties='Excel 12.0 Xml;HDR=YES'";

            string query = "SELECT * FROM ["+HojaExcel+"$]";


            //Creamos el provider
            OleDbConnection excelConnection = new OleDbConnection(connectionString);
            //Lo abrimos
            excelConnection.Open();
            //Creamos un Data Adapter que extraiga los datos necesarios(todos) del provider
            OleDbDataAdapter data = new OleDbDataAdapter(query, excelConnection);
            //Creamos una tabla
            DataTable dTable = new DataTable();
            //Usando el Data Adapter que tiene los datos seleccionados, rellenamos la tabla.
            data.Fill(dTable);

            return dTable;
        }

        public static void CargueABaseDeDatos(DataTable dt, string nombreArchivo, string tablaDestino, string tablaControl, string procedimientoFinal){

            int control=0;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.Text;
            
            using (SqlConnection connection =
                   new SqlConnection(""))
            {
                cmd.Connection = connection;
                connection.Open();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = tablaDestino;

                    
                    try
                    {
                        // Write from the source to the destination.
                        cmd.CommandText = "Select count(*) from " + tablaControl + " where Archivo='" + nombreArchivo + "'";
                        control = (int)cmd.ExecuteScalar();

                        if (control==0)
                        {
                            bulkCopy.WriteToServer(dt);
                            cmd.CommandText = "Insert into " + tablaControl + " values('" + nombreArchivo + "','" + DateTime.Now.ToString() + "')";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = "Exec "+procedimientoFinal+"";
                            cmd.ExecuteNonQuery();
                        }

                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
        }


    }
}

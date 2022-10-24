
using SpreadsheetLight;
using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.IO;
using System;
using System.Reflection.Metadata.Ecma335;
using System.Text.Json.Serialization;
using Newtonsoft.Json;


namespace ProgramaJson
{
    public class Program
    {

       
        string rutaArchivo;
        int tamanoFilas;
        int tamanoColumnas;

        

        static void Main(string[] args)
        {
            

            var ente = GetEnte();
            SerializeJsonFile(ente);
            
        }

        public static void SerializeJsonFile(List<Ente> listaEnte)
        {
            try
            {
                string rutaDestino = @"C:/Prueba/" + "Declaranet -" + DateTime.Now.ToString("dd-MM-yyyy hhmmss") + ".json";
                string generarJson = JsonConvert.SerializeObject(listaEnte.ToArray(), Formatting.Indented);
                
                File.WriteAllText(rutaDestino, generarJson);
            }
            catch (Exception exp)
            {

                Console.WriteLine(exp.ToString());
            }


        }


        public static List<Ente> GetEnte()
        {
            
            string rutaArchivo;
            Console.WriteLine("Ingrese la ruta del archivo");
            rutaArchivo = @"C:\Prueba\real.xlsx";

            SLDocument sl = new SLDocument(rutaArchivo);
            SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics();
            
            

            int tamanoFila = propiedades.EndRowIndex; 
            int tamanoColumna = propiedades.EndColumnIndex; 


            List<Ente> listaEnte = new List<Ente>();

            for (int i = 2; i <= tamanoFila; i++)
            {
                var objEnte = new Ente();
                
                for (int j = 1; j < tamanoColumna ;)
                {
                    

                    var nivelGobierno = new NivelGobierno();
                    var datosDeControl = new DatosDeControl();
                    var entidadFederativa = new EntidadFederativa();

                    objEnte.Id = "";
                    objEnte.EnteDes = sl.GetCellValueAsString(i, j);
                    
                    objEnte.IdentificadorINAI = 1;

                    objEnte.NombreCorto = sl.GetCellValueAsString(i, ++j);

                    nivelGobierno.Nombre = sl.GetCellValueAsString(i, ++j);

                    objEnte.Poder = sl.GetCellValueAsString(i, ++j);
                    objEnte.SecNac = "No_Se_Nac";
                    objEnte.TipoEntidad = sl.GetCellValueAsString(i, ++j);
                    objEnte.UnidadResponsable = "0";
                    objEnte.idEnteOrigen = "";

                    objEnte.NivelGobierno = nivelGobierno;
                    nivelGobierno.EntidadFederativa = entidadFederativa;
                    entidadFederativa.entidadFederativaDesc = sl.GetCellValueAsString(i, ++j);
                    entidadFederativa.idEntidadFederativa = 4;

                    objEnte.Ramo = sl.GetCellValueAsInt32(i, ++j);
                    objEnte.Rfc = sl.GetCellValueAsString(i, ++j);

             
                    datosDeControl.UsuarioActualiza = 0;
                    datosDeControl.Situacion = "NORMAL";
                    datosDeControl.FechaRegistro = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss");
                    datosDeControl.FechaUltimaActualiza = DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss");
                    datosDeControl.UsuarioRegistra = 0;
                    datosDeControl.Activo = 1;
                    
                    objEnte.DatosDeControl = datosDeControl;
                    
                    

                }

                listaEnte.Add(objEnte);
                Console.WriteLine(listaEnte.ToString());


            }

            
            return listaEnte;
        }
        


    }
}



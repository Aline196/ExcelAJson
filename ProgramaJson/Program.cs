
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

        static void Main(String[] args)
        {

           var ente = GetEnte();

        }

        public static List<Ente> GetEnte()
        {
            string rutaArchivo;
            Console.WriteLine("Ingrese la ruta del archivo");
            rutaArchivo = Console.ReadLine();

            SLDocument sl = new SLDocument(rutaArchivo);
            SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics();

            int tamanoFila = propiedades.EndRowIndex; //10
            int tamanoColumna = propiedades.EndColumnIndex; //6


            List<Ente> listaEnte = new List<Ente>();

            //while (!string.IsNullOrEmpty())   
            for (int i = 2; i <= tamanoFila; i++)
            {
                 var objEnte = new Ente();
                

                for (int j = 1; j < tamanoColumna;)
                {
                    var nivelGobierno = new NivelGobierno();
                    var entidadFederativa = new EntidadFederativa();
        
                    objEnte.EnteDes = sl.GetCellValueAsString(i, j);
                    objEnte.NombreCorto = sl.GetCellValueAsString(i, ++j);

                    nivelGobierno.Nombre = sl.GetCellValueAsString(i, ++j);


                    objEnte.Poder = sl.GetCellValueAsString(i, ++j);

                    entidadFederativa.entidadFederativaDesc = sl.GetCellValueAsString(i, ++j);
                    nivelGobierno.EntidadFederativa = entidadFederativa;
                    objEnte.NivelGobierno = nivelGobierno;

                    objEnte.Ramo = sl.GetCellValueAsDouble(i, ++j);
                    objEnte.Rfc = sl.GetCellValueAsString(i, ++j);

                   






                }
                
                listaEnte.Add(objEnte);

            }

            return listaEnte;
        }

        public static void serializarJson(List<Ente> ent)
        {
            //string obtenerJson = JsonConverter(ent.ToArray(), Formatting.Indented);

            //JsonConverter.Ser









        }
        


    }
}



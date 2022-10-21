using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProgramaJson
{
    public class Ente
    {
        public string Id { get; set; }
        public DatosDeControl? datosDeControl { get; set; }
        public string EnteDes { get; set; }
        public double IdentificadorINAI { get; set; }
        public NivelGobierno? NivelGobierno { get; set; }
        public string NombreCorto { get; set; }
        public string Poder { get; set; }
        public double Ramo { get; set; }
        public string Rfc { get; set; }
        public string SecNac { get; set; }


    }
}

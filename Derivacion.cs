using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeneradorDeMensajes
{
    internal class Derivacion
    {

        private String rolOficio;
        private String partes;
        private String isapre;
        private String tribunal;
        private String formaDeIngreso;
        private String materia;
        private String fechaDeDerivacion;
        private String fechaDeAudienciaReal;
        private String asignado;
        private String fechaDeAntecedentes;
        private String antecedentesEnviados;
        private String pjud;
        private String folio;
        private String direccionDeCorreo;

        public Derivacion()
        {
        }

        public Derivacion(string rolOficio, string partes, string isapre, string tribunal, string formaDeIngreso, string materia, string fechaDeDerivacion, string fechaDeAudienciaReal, string asignado, string fechaDeAntecedentes, string antecedentesEnviados, string pjud, string folio, string direccionDeCorreo)
        {
            this.RolOficio = rolOficio;
            this.Partes = partes;
            this.Isapre = isapre;
            this.Tribunal = tribunal;
            this.FormaDeIngreso = formaDeIngreso;
            this.Materia = materia;
            this.FechaDeDerivacion = fechaDeDerivacion;
            this.FechaDeAudienciaReal = fechaDeAudienciaReal;
            this.Asignado = asignado;
            this.FechaDeAntecedentes = fechaDeAntecedentes;
            this.AntecedentesEnviados = antecedentesEnviados;
            this.Pjud = pjud;
            this.Folio = folio;
            this.DireccionDeCorreo = direccionDeCorreo;
        }

        public string RolOficio { get => rolOficio; set => rolOficio = value; }
        public string Partes { get => partes; set => partes = value; }
        public string Isapre { get => isapre; set => isapre = value; }
        public string Tribunal { get => tribunal; set => tribunal = value; }
        public string FormaDeIngreso { get => formaDeIngreso; set => formaDeIngreso = value; }
        public string Materia { get => materia; set => materia = value; }
        public string FechaDeDerivacion { get => fechaDeDerivacion; set => fechaDeDerivacion = value; }
        public string FechaDeAudienciaReal { get => fechaDeAudienciaReal; set => fechaDeAudienciaReal = value; }
        public string Asignado { get => asignado; set => asignado = value; }
        public string FechaDeAntecedentes { get => fechaDeAntecedentes; set => fechaDeAntecedentes = value; }
        public string AntecedentesEnviados { get => antecedentesEnviados; set => antecedentesEnviados = value; }
        public string Pjud { get => pjud; set => pjud = value; }
        public string Folio { get => folio; set => folio = value; }
        public string DireccionDeCorreo { get => direccionDeCorreo; set => direccionDeCorreo = value; }
    }
}

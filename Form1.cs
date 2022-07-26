using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Storage;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Font = Xceed.Document.NET.Font;
using LicenseContext = OfficeOpenXml.LicenseContext;


namespace GeneradorDeMensajes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //programa debe tomar Excel de una hoja y generar varios words a partir de la información presente en dicho Excel

            //actualizacion 14/07/2022
            //El programa debe hacer 2 cosas: primero generar un word de todas las demandas que se derivaron el día actual (si es que
            //no estan generados ya). Además, debe generar un reccordatorio con todas las demandas la cuya audiencia real sea la semana
            //siguiente (independiente de si el día en cuestión es no hábil). 
        }

        private void button1_Click(object sender, EventArgs e)
        {

            MessageBox.Show("Seleccione Excel (una vez que se termine de editar, por el día)");

            int recordatoriosTotales = 0;

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            while (true)
            {
                if (choofdlog.ShowDialog() == DialogResult.OK)
                {
                    sFileName = choofdlog.FileName;
                    arrAllFiles = choofdlog.FileNames; //used when Multiselect = true
                    break;
                }
                else
                {
                    MessageBox.Show("No se seleccionó nada, proceso terminado.");
                    System.Environment.Exit(0);
                    
                }
            }


            try
            {

         

            List<Derivacion> derivaciones = leerExcelDeDerivaciones(sFileName);

            var hoySinTiempo = DateTime.Now;



            string[] hoySinTiempoArray = hoySinTiempo.ToString().Split(' ');
            String hoySinTiempoComoString = hoySinTiempoArray[0];
            
            hoySinTiempoComoString = hoySinTiempoComoString.Replace("/", "-");
            Console.WriteLine(hoySinTiempoComoString);

            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads\MensajesGenerados"; //el día "+hoySinTiempoComoString+ @" \" ;

            DirectoryInfo di = Directory.CreateDirectory(downloads);

            int contadorDeWordsCreadosConFechaDeDerivacionActual = 0;

            List<String> asignadasUnicas = new List<String>();
            foreach (var item in derivaciones)
            {
                if (item.Asignado!="Asignado")
                {
                    asignadasUnicas.Add(item.Asignado);
                }
                
            }

            asignadasUnicas=asignadasUnicas.ToArray().Distinct().ToList();

            foreach (var asignada in asignadasUnicas)//personas asignadas
            {


                var hoy = DateTime.Now;

                String fechaComoString = "";

                    string frase = hoy.ToString();
                    string[] palabras = frase.Split(' ');
                    fechaComoString = palabras[0];
                
                fechaComoString = fechaComoString.Replace("/","-");

                    String archivo = downloads + @"\Demandas asignadas a " + asignada +" el "+fechaComoString+ ".docx";

                if (File.Exists(archivo))//archivo existe, no se crea nada
                {

                }
                else//archivo no existe, así que se crea
                {//deberia agregarse tambien que si no hay demandas para esta persona, entonces no se crea


                    var doc = DocX.Create(archivo);

                    //titulo que va a llevar el word
                    string title = "Estimada " + asignada + "," + Environment.NewLine;

                    //texto a escribir en el word
                    string textParagraph = "El día de hoy se le han asignado nuevos oficios correspondientes" +
                        " a las siguientes causas:" + Environment.NewLine + Environment.NewLine;

                    String informacionDeDemanda = "";


                    foreach (var item in derivaciones)//posibles demandas
                    {

                        //obtener fecha actual
                        var j = DateTime.Now;

                        String soloLaFechaActualComoString = "";
                        String soloLaFechaDeDerivacionComoString = "";
                        String soloLaFechaDeAudienciaRealComoString = "";
                        String soloLaFechaDeAntecedentesComoString = "";
                        String soloLaFechaPjudComoString = "";

                        //quitar el tiempo y mantener la fecha
                        if (item.FechaDeDerivacion != null)
                        {

                            string phrase = j.ToString();
                            string[] words = phrase.Split(' ');
                            soloLaFechaActualComoString = words[0];

                            phrase = item.FechaDeDerivacion.ToString();
                            words = phrase.Split(' ');
                            soloLaFechaDeDerivacionComoString = words[0];
                            item.FechaDeDerivacion = soloLaFechaDeDerivacionComoString;


                            phrase = item.FechaDeAudienciaReal.ToString();
                            words = phrase.Split(' ');
                            soloLaFechaDeAudienciaRealComoString = words[0];
                            item.FechaDeAudienciaReal = soloLaFechaDeAudienciaRealComoString;

                            //phrase = item.FechaDeAntecedentes.ToString();
                            //words = phrase.Split(' ');
                            //soloLaFechaDeAntecedentesComoString = words[0];
                            //item.FechaDeAntecedentes = soloLaFechaDeAntecedentesComoString;

                            phrase = item.Pjud.ToString();
                            words = phrase.Split(' ');
                            soloLaFechaPjudComoString = words[0];
                            item.Pjud = soloLaFechaPjudComoString;

                        }


                        if ((item.Asignado == asignada) && (item.RolOficio != "Rol Oficio" && item.RolOficio != null) && (soloLaFechaActualComoString == soloLaFechaDeDerivacionComoString))
                        {
                                                  
                            informacionDeDemanda = "-	" + item.Tribunal + ", Rol " + item.RolOficio + ", caratulada “" + validarPartes(item.Partes) + "” respecto de Isapre " + item.Isapre + ", " +
                           "con fecha de audiencia para el día " + convertirFechaAPalabras(item.FechaDeAudienciaReal) + "." + Environment.NewLine + Environment.NewLine;

                            textParagraph += informacionDeDemanda;
                        }

                        

                    }

                    Formatting titleFormat = new Formatting();
                    //Specify font family  
                    titleFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    //Specify font size  
                    titleFormat.Size = 12;
                    // titleFormat.Position = 40;
                    titleFormat.FontColor = System.Drawing.Color.Black;
                    //titleFormat.UnderlineColor = System.Drawing.Color.Gray;
                    //titleFormat.Italic = true;

                    //Formatting Text Paragraph  
                    Formatting textParagraphFormat = new Formatting();
                    //font family  
                    textParagraphFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    //font size  
                    textParagraphFormat.Size = 12;
                    //Spaces between characters  
                    // textParagraphFormat.Spacing = 2;


                    doc.InsertParagraph(title, false, titleFormat);
                    doc.InsertParagraph(textParagraph, false, titleFormat);//textParagraphFormat

                    doc.Save();
                    contadorDeWordsCreadosConFechaDeDerivacionActual++;
                }

            }


           
            //si el día en el que corre el proceso, es un jueves, entonces deben crearse words de recordatorio de 
            //la semana que viene (para enviarse ese mismo jueves, o sea, la fecha actual).
            
            var x = DateTime.Now;
            if (x.DayOfWeek == DayOfWeek.Thursday)
            {
                Console.WriteLine("Es Jueves");
                MessageBox.Show("Es jueves; se generará un recordatorio");
                //deben crearse un documento extra con todas las demandas (separadas por abogada asignada)
                //actualizacion 18/07/2022: debe crearse UN SOLO documento de recordatorio para todas.
        
              
                    String proximasDemandas ="";
                    int contadorDeDemandasDeLaSemanaQueViene = 0;

                    foreach (var item2 in derivaciones)
                    {

                        var j = DateTime.Now;
                        String proximoLunes=j.AddDays(3).ToString();
                        var proximoMartes = j.AddDays(4).ToString();
                        var proximoMiercoles = j.AddDays(5).ToString();
                        var proximoJueves = j.AddDays(6).ToString();
                        var proximoViernes = j.AddDays(7).ToString();
                        var proximoSabado = j.AddDays(8).ToString();
                        var proximoDomingo = j.AddDays(9).ToString();


                        string[] words = proximoLunes.Split(' ');
                        String proximoLunesSinTiempo = words[0];

                         words = proximoMartes.Split(' ');
                        String proximoMartesSinTiempo = words[0];

                         words = proximoMiercoles.Split(' ');
                        String proximoMiercolesSinTiempo = words[0];

                         words = proximoJueves.Split(' ');
                        String proximoJuevesSinTiempo = words[0];

                         words = proximoViernes.Split(' ');
                        String proximoViernesSinTiempo = words[0];

                        words = proximoSabado.Split(' ');
                        String proximoSabadoSinTiempo = words[0];

                        words = proximoDomingo.Split(' ');
                        String proximoDomingoSinTiempo = words[0];

                        List<String> proximasFechas=new List<String>();
                        proximasFechas.Add(proximoLunesSinTiempo);
                        proximasFechas.Add(proximoMartesSinTiempo);
                        proximasFechas.Add(proximoMiercolesSinTiempo);
                        proximasFechas.Add(proximoJuevesSinTiempo);
                        proximasFechas.Add(proximoViernesSinTiempo);
                        proximasFechas.Add(proximoSabadoSinTiempo);
                        proximasFechas.Add(proximoDomingoSinTiempo);

                        String[] listadoDeProximasFechas = proximasFechas.ToArray();

                        if (listadoDeProximasFechas.Contains(item2.FechaDeAudienciaReal))//si hay unda demanda la semana que viene
                        {
                            proximasDemandas += "-	Rol " + item2.RolOficio + ", fecha de audiencia para el día " + convertirFechaAPalabras(item2.FechaDeAudienciaReal) + "." + Environment.NewLine;


                            contadorDeDemandasDeLaSemanaQueViene++;
                        }
                    }

                var fechaHoy = DateTime.Now;
                String proximoLunesComoString = fechaHoy.AddDays(4).ToString();
                string[] division = proximoLunesComoString.Split(' ');
                String proximoLunesSinTiempoComoString = division[0];
                Console.WriteLine(proximoLunesSinTiempoComoString);
                string[] proximoJuevesSeparado = proximoLunesSinTiempoComoString.Split('/');
       


                String dia = proximoJuevesSeparado[1] + " de ";
                String mesComoPalabra = "";
                switch (proximoJuevesSeparado[0])
                {
                    case "01":
                        mesComoPalabra = "enero";
                        break;
                    case "02":
                        mesComoPalabra = "febrero";
                        break;
                    case "03":
                        mesComoPalabra = "marzo";
                        break;
                    case "04":
                        mesComoPalabra = "abril";
                        break;
                    case "05":
                        mesComoPalabra = "mayo";
                        break;
                    case "06":
                        mesComoPalabra = "junio";
                        break;
                    case "07":
                        mesComoPalabra = "julio";
                        break;
                    case "08":
                        mesComoPalabra = "agosto";
                        break;
                    case "09":
                        mesComoPalabra = "septiembre";
                        break;
                    case "10":
                        mesComoPalabra = "octubre";
                        break;
                    case "11":
                        mesComoPalabra = "noviembre";
                        break;
                    case "12":
                        mesComoPalabra = "diciembre";
                        break;                      
                    default:
                        mesComoPalabra = "";
                        break;
                }


                String mes = mesComoPalabra+" del ";
                String anio = "20"+ proximoJuevesSeparado[2];

                String lunesDeLaSemanaDeRecordatorios =dia+mes+anio;
                    //crear word de recordatorio

                    String archivo = downloads + @"\Recordatorio oficios de semana siguiente ("+ lunesDeLaSemanaDeRecordatorios + @").docx";


                    if (File.Exists(archivo))//archivo existe, no se crea nada
                    {

                    }
                    else//archivo no existe, asi que se crea
                    {

                        var doc = DocX.Create(archivo);

                        //titulo que va a llevar el word
                        string title = "Estimadas, " + Environment.NewLine;

                        //texto a escribir en el word
                        string textParagraph = "Para la próxima semana, tenemos los siguientes oficios con audiencia: "+Environment.NewLine+Environment.NewLine+
                        proximasDemandas;

                    Formatting titleFormat = new Formatting();                   
                    titleFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    titleFormat.Size = 12;                  
                    titleFormat.FontColor = System.Drawing.Color.Black;             
                   
                    Formatting textParagraphFormat = new Formatting();                     
                    textParagraphFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");                   
                    textParagraphFormat.Size = 12;

                    doc.InsertParagraph(title, false, titleFormat);
                    doc.InsertParagraph(textParagraph, false, titleFormat);//textParagraphFormat

                        if (contadorDeDemandasDeLaSemanaQueViene > 0)
                        {
                            doc.Save();

                        recordatoriosTotales++;

                        }
                    }
    
            }
            else
            {
                Console.WriteLine("No es jueves");
                //no se crean documentos extras
                
            }


            contadorDeWordsCreadosConFechaDeDerivacionActual += recordatoriosTotales;



            if (contadorDeWordsCreadosConFechaDeDerivacionActual > 1)
            {
                MessageBox.Show("Se crearon "+contadorDeWordsCreadosConFechaDeDerivacionActual.ToString()+" documentos nuevos en la carpeta de descargas para enviar hoy");
            }else if (contadorDeWordsCreadosConFechaDeDerivacionActual==1)
            {
                MessageBox.Show("Se creó un documento nuevo en la carpeta de descargas para enviar hoy");
            }else if (contadorDeWordsCreadosConFechaDeDerivacionActual==0)
            {
                MessageBox.Show("No se creó ningún documento para enviar hoy");
            }

            }
            catch (Exception)
            {
                MessageBox.Show("Debe seleccionar un archivo válido.");
                MessageBox.Show("Excel a subir debe tener las siguientes columnas:" +
                    " A: Rol Oficio, " +
                    " B: Partes, " +
                    " C: Isapre, " +
                    " D: Tribunal, " +
                    " E: Forma de ingreso, " +
                    " F: Materia, " +
                    " G: Fecha de derivacion, "+
                    " H: Fecha de audiencia, " +
                    " I: Asignado, " +
                    " J: Pjud, " +
                    " K: Folio " );
                MessageBox.Show("Cerrando aplicación.");
                System.Environment.Exit(0);
                throw;
            }

        }

        private List<Derivacion> leerExcelDeDerivaciones(String sFileName)
        {
            List<Derivacion> listadoDeDerivaciones = new List<Derivacion>();

      
            FileInfo existingFile = new FileInfo(sFileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                int cantidadDeHojas = package.Workbook.Worksheets.Count;


                for (int i = 0; i < cantidadDeHojas; i++)
                {

                    //get the first worksheet in the workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                    int colCount = worksheet.Dimension.End.Column;  //get Column Count
                    int rowCount = worksheet.Dimension.End.Row;     //get row count

                    for (int row = 1; row <= rowCount; row++)
                    {
                        Derivacion d = new Derivacion();

                        d.RolOficio = worksheet.Cells[row, 1].Value?.ToString().Trim();
                        d.Partes = worksheet.Cells[row, 2].Value?.ToString().Trim();
                        d.Isapre = worksheet.Cells[row, 3].Value?.ToString().Trim();
                        d.Tribunal = worksheet.Cells[row, 4].Value?.ToString().Trim();
                        d.FormaDeIngreso = worksheet.Cells[row, 5].Value?.ToString().Trim();
                        d.Materia = worksheet.Cells[row, 6].Value?.ToString().Trim();
                        d.FechaDeDerivacion = worksheet.Cells[row, 7].Value?.ToString().Trim();
                        d.FechaDeAudienciaReal = worksheet.Cells[row, 8].Value?.ToString().Trim();
                        d.Asignado = worksheet.Cells[row, 9].Value?.ToString().Trim();
                        //d.FechaDeAntecedentes = worksheet.Cells[row, 10].Value?.ToString().Trim(); //actualizacion 18/07/2022--> no necesita esta informacion en el word
                        //d.AntecedentesEnviados = worksheet.Cells[row, 11].Value?.ToString().Trim(); //actualizacion 18/07/2022--> no necesita esta informacion en el word
                        d.Pjud = worksheet.Cells[row, 10].Value?.ToString().Trim();
                        d.Folio = worksheet.Cells[row, 11].Value?.ToString().Trim();
                        //d.DireccionDeCorreo = worksheet.Cells[row, 14].Value?.ToString().Trim(); //actualizacion 18/07/2022--> no necesita esta informacion en el word

                        listadoDeDerivaciones.Add(d);

                    }
                }

            }
           
            return listadoDeDerivaciones;

        }

        private String convertirFechaAPalabras(String fechaEnFormatoFecha)//mes, dia, anio
        {
            String fechaComoPalabras = "";

            string[] palabras = fechaEnFormatoFecha.Split('/');

            //fecha viene asi 07/18/2022

            String dia = palabras[1];
            String mes = palabras[0];
            String anio = palabras[2];

            switch (mes)
            {
                case "01":
                    mes = "enero";
                    break;
                case "02":
                    mes = "febrero";
                    break;
                case "03":
                    mes = "marzo";
                    break;
                case "04":
                    mes = "abril";
                    break;
                case "05":
                    mes = "mayo";
                    break;
                case "06":
                    mes = "junio";
                    break;
                case "07":
                    mes = "julio";
                    break;
                case "08":
                    mes = "agosto";
                    break;
                case "09":
                    mes = "septiembre";
                    break;
                case "10":
                    mes = "octubre";
                    break;
                case "11":
                    mes = "noviembre";
                    break;
                case "12":
                    mes = "diciembre";
                    break;
                default:
                    
                    break;

            }

            fechaComoPalabras = dia+" de "+mes+" de 20"+anio+"";

            return fechaComoPalabras;
        }


        private String validarPartes(String partes)
        {
            String partesValidadas = partes;

            if (String.IsNullOrEmpty(partesValidadas))
            {
                partesValidadas = "Partes no detalladas";
            }

            return partesValidadas;
        }


    }
}

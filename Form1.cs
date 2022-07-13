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
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string sFileName = "";

            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Archivos XLSX (*.xlsx)|*.xlsx";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = true;

            string[] arrAllFiles = new string[] { };

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                sFileName = choofdlog.FileName;
                arrAllFiles = choofdlog.FileNames; //used when Multiselect = true           
            }

            List<Derivacion> derivaciones = leerExcelDeDerivaciones(sFileName);





            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            int contadorDeWordsCreados = 0;




            foreach (var item in derivaciones)
            {

                //obtener fecha actual
                var j = DateTime.Now;

                String soloLaFechaActualComoString = "";
                String soloLaFechaDeDerivacionComoString = "";
                String soloLaFechaDeAudienciaRealComoString = "";
                String soloLaFechaDeAntecedentesComoString = "";
                String soloLaFechaPjudComoString = "";


                //quitar el tiempo y mantener la fecha
                if (item.FechaDeDerivacion!=null)
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

                    phrase = item.FechaDeAntecedentes.ToString();
                    words = phrase.Split(' ');
                    soloLaFechaDeAntecedentesComoString = words[0];
                    item.FechaDeAntecedentes = soloLaFechaDeAntecedentesComoString;

                    phrase = item.Pjud.ToString();
                    words = phrase.Split(' ');
                    soloLaFechaPjudComoString = words[0];
                    item.Pjud = soloLaFechaPjudComoString;

                }
                

                //si la fecha actual es igual a la fecha de derivacion, entonces se crea el word, para hoy
                if ((item.RolOficio!= "Rol Oficio" && item.RolOficio != null) && (soloLaFechaActualComoString==soloLaFechaDeDerivacionComoString))
                {
                    String archivo = downloads + @"\Rol oficio " + item.RolOficio + ".docx";

                    var doc = DocX.Create(archivo);

                    //titulo que va a llevar el word
                    string title = "Estimado/a " + item.Asignado + ",";

                    //texto a escribir en el word
                    string textParagraph = "usted a sido asignado/a para el " + item.Tribunal + " el día " + item.FechaDeAudienciaReal + " por el rol de oficio " + item.RolOficio + " de la Isapre"
                        + item.Isapre + ", entre las partes " + item.Partes + ". En materia de " + item.Materia + ". " + Environment.NewLine
                        + "La fecha de derivacion fue el " + item.FechaDeDerivacion + ". La fecha en que los antecedentes fueron enviados fue el " + item.FechaDeAntecedentes + " y se encuentran en estado" + item.AntecedentesEnviados
                        + ". El pjud es " + item.Pjud + " y el folio está " + item.Folio;



                    Formatting titleFormat = new Formatting();
                    //Specify font family  
                    titleFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    //Specify font size  
                    titleFormat.Size = 12;
                    titleFormat.Position = 40;
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
                    textParagraphFormat.Spacing = 2;


                    doc.InsertParagraph(title, false, titleFormat);
                    doc.InsertParagraph(textParagraph, false, titleFormat);//textParagraphFormat


                    doc.Save();
                    contadorDeWordsCreados++;


                }


            }



            var x = DateTime.Now;
            if (x.DayOfWeek != DayOfWeek.Friday)
            {
                Console.WriteLine("Es viernes");
                //deben crearse documentos extra
                List<string> listadoDeAbogados = new List<string>();
                foreach (var item in derivaciones)
                {
                    listadoDeAbogados.Add(item.Asignado);
                }

                List<String> listadoDeAbogadosConDuplicados = listadoDeAbogados;
                List<String> listadoDeAbogadosSinDuplicados = listadoDeAbogadosConDuplicados.Distinct().ToList();


                foreach (var item in listadoDeAbogadosSinDuplicados)
                {
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

                        List<String> proximasFechas=new List<String>();
                        proximasFechas.Add(proximoLunesSinTiempo);
                        proximasFechas.Add(proximoMartesSinTiempo);
                        proximasFechas.Add(proximoMiercolesSinTiempo);
                        proximasFechas.Add(proximoJuevesSinTiempo);
                        proximasFechas.Add(proximoViernesSinTiempo);

                        String[] listadoDeProximasFechas = proximasFechas.ToArray();

                        if (item.ToString()==item2.Asignado && listadoDeProximasFechas.Contains(item2.FechaDeAudienciaReal))//si el registro
                        {
                            proximasDemandas += "el día "+item2.FechaDeAudienciaReal+" "+Environment.NewLine;
                            contadorDeDemandasDeLaSemanaQueViene++;
                        }
                    }

                    proximasDemandas += ".";

                    //crear word de recordatorio

                    String archivo = downloads + @"\Recordatorio para " + item.ToString() + ".docx";

                    var doc = DocX.Create(archivo);

                    //titulo que va a llevar el word
                    string title = "Estimado/a " + item.ToString() + ",";

                    //texto a escribir en el word
                    string textParagraph = "se le recuerda que tiene demandas asignadas para la semana que viene. Especificamente los días:"+Environment.NewLine
                        +proximasDemandas;



                    Formatting titleFormat = new Formatting();
                    //Specify font family  
                    titleFormat.FontFamily = new Xceed.Document.NET.Font("Times New Roman");
                    //Specify font size  
                    titleFormat.Size = 12;
                    titleFormat.Position = 40;
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
                    textParagraphFormat.Spacing = 2;


                    doc.InsertParagraph(title, false, titleFormat);
                    doc.InsertParagraph(textParagraph, false, titleFormat);//textParagraphFormat

                    if (contadorDeDemandasDeLaSemanaQueViene>0)
                    {
                        doc.Save();
                    }
                    


                }

            }
            else
            {
                Console.WriteLine("No es viernes");
                //no se crean documentos extras
                
            }



            if (contadorDeWordsCreados > 1)
            {
                MessageBox.Show("Se crearon "+contadorDeWordsCreados.ToString()+" documentos en la carpeta de descargas para enviar hoy");
            }else if (contadorDeWordsCreados==1)
            {
                MessageBox.Show("Se creó un documento en la carpeta de descargas para enviar hoy");
            }else if (contadorDeWordsCreados==0)
            {
                MessageBox.Show("No se creó ningún documento para enviar hoy");
            }
            



        }

        private List<Derivacion> leerExcelDeDerivaciones(String sFileName)
        {
            List<Derivacion> listadoDeDerivaciones = new List<Derivacion>();

      
            FileInfo existingFile = new FileInfo(sFileName);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count


                for (int row = 1; row <= rowCount; row++)
                {

                    Derivacion d =  new Derivacion();   

                    d.RolOficio = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    d.Partes = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    d.Isapre = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    d.Tribunal = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    d.FormaDeIngreso = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    d.Materia = worksheet.Cells[row, 6].Value?.ToString().Trim();
                    d.FechaDeDerivacion = worksheet.Cells[row, 7].Value?.ToString().Trim();
                    d.FechaDeAudienciaReal = worksheet.Cells[row, 8].Value?.ToString().Trim();
                    d.Asignado = worksheet.Cells[row, 9].Value?.ToString().Trim();
                    d.FechaDeAntecedentes = worksheet.Cells[row, 10].Value?.ToString().Trim();
                    d.AntecedentesEnviados = worksheet.Cells[row, 11].Value?.ToString().Trim();
                    d.Pjud = worksheet.Cells[row, 12].Value?.ToString().Trim();
                    d.Folio = worksheet.Cells[row, 13].Value?.ToString().Trim();
                    d.DireccionDeCorreo = worksheet.Cells[row, 14].Value?.ToString().Trim();



                    listadoDeDerivaciones.Add(d);   

                }





            }
            


            return listadoDeDerivaciones;

        }
    }
}

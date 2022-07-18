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
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace FiltradorDePlanillas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

            List<Registro> registros = leerExcelDePlanillasDeLaCCU(sFileName);


            string localfolder = ApplicationData.Current.LocalFolder.Path;
            var array = localfolder.Split('\\');
            var username = array[2];
            string downloads = @"C:\Users\" + username + @"\Downloads";


            var archivo = new FileInfo(downloads + @"\RegistrosFiltrados.xlsx");

            SaveExcelFileAusencia(registros, archivo);

            MessageBox.Show("Archivo Excel filtrado, creado en carpeta de descargas!");


        }




        private List<Registro> leerExcelDePlanillasDeLaCCU( String filePath){
            List<Registro> registros = new List<Registro>();
            List<String> planillas = new List<String>();

            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count



                for (int row = 1; row <= rowCount; row++)
                {

                    Registro r = new Registro();
                    r.Uen = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    r.Cd = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    r.CentroDeDistribucion = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    r.Fletero = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    r.Nombre = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    r.Camion = worksheet.Cells[row, 6].Value?.ToString().Trim();
                    r.SaldoAnterior = worksheet.Cells[row, 7].Value?.ToString().Trim();
                    r.Planilla = worksheet.Cells[row, 8].Value?.ToString().Trim();
                    r.ValoresAEntregar = worksheet.Cells[row, 9].Value?.ToString().Trim();
                    r.ValoresEEntregados = worksheet.Cells[row, 10].Value?.ToString().Trim();
                    r.SaldoCredito = worksheet.Cells[row, 11].Value?.ToString().Trim();
                    r.SaldoDebito = worksheet.Cells[row, 12].Value?.ToString().Trim();
                    r.Diferencia = worksheet.Cells[row, 13].Value?.ToString().Trim();
                    r.FechaPlanilla = worksheet.Cells[row, 14].Value?.ToString().Trim();
                    r.FechaCierre = worksheet.Cells[row, 15].Value?.ToString().Trim();
                    r.Observaciones = worksheet.Cells[row, 16].Value?.ToString().Trim();
                    r.Referencia = worksheet.Cells[row, 17].Value?.ToString().Trim();

                    String observacion = r.Observaciones;

                    if (observacion != null)
                    {

                    string[] words = observacion.Split(':');
                    

                    //todo registro con una observacion que empiece con "Carga de Reparto", debe ser filtrada
                    //hasta aqui esta CORRECTO, falta identificar las duplicadas y luego cambiarles el número
                    if (words[0]!= "Carga de Reparto")
                    {

                        registros.Add(r);
                        planillas.Add(r.Planilla);
                    }

                    }
                }

            }

            //las planillas unicas
            planillas = planillas.Distinct().ToList();

            List<Registro> planillasDuplicadas = new List<Registro>();

            foreach (var item in registros)
            {

                int contadorDePlanillas = 0;
                foreach (var item2 in planillas)
                {
                   

                    if (item2==item.Planilla)
                    {
                        
                        contadorDePlanillas++;
                        if (contadorDePlanillas > 1)
                        {
                            planillasDuplicadas.Add(item);
                        }
                    }



                }

          


            }
           



            //String hoy = DateTime.Now.ToString("yyyy.MM.MM");
            //hoy +=01.ToString();

            ////obtener fecha actual
            //foreach (var item in registros)
            //{

            //}

         //sacar carga de reparto, despues de identificar las duplicadas

            return registros;

        }




        private static async Task SaveExcelFileAusencia(List<Registro> registros, FileInfo file)
        {
            var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Registros filtrados");

            var range = ws.Cells["A1"].LoadFromCollection(registros, true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }




    }
}

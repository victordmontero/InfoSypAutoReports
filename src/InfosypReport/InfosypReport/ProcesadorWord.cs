using System;
using System.Linq;
using System.IO;
using Novacode;
using System.Drawing;

namespace InfosypReport
{
    public class ProcesadorWord : IDisposable
    {
        public const float ANCHO = 100.0f;
        public DocX Documento { get; set; }

        public ProcesadorWord(string ArchivoPlantilla)
        {
            AbrirDoc(ArchivoPlantilla);
        }

        public void AbrirDoc(string ArchivoPlantilla)
        {
            try
            {
                if (File.Exists(ArchivoPlantilla))
                {
                    Documento = DocX.Load(ArchivoPlantilla);
                }
                else
                {
                    throw new ArgumentException("El archivo no existe en la carpeta temporal.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generado. Descripción: " + ex.Message);
            }
        }

        public void GuardarComo(string ruta)
        {
            try
            {
                string path = Path.ChangeExtension(ruta, ".docx");
                Documento.SaveAs(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al guardar: " + ex.Message);
            }
        }

        public void InsertarEntrada(string texto, Formatting formato)
        {
            Documento.InsertParagraph(texto, false, formato).SpacingBefore(2).SpacingAfter(10);
        }

        public void InsertarFechaCabezera(DateTime fecha, double horas)
        {
            var tabla = Documento.AddTable(1, 2);

            var parrafo1 = tabla.Rows[0].Cells[0].Paragraphs.First();
            parrafo1.Append(fecha.ToString("dd/MM/yyyy"));
            parrafo1.Alignment = Alignment.left;
            parrafo1.Bold();
            parrafo1.FontSize(8);
            parrafo1.Font(new FontFamily("Arial"));

            var parrafo2 = tabla.Rows[0].Cells[1].Paragraphs.First();
            parrafo2.Append(string.Format("{0}", TimeSpan.FromHours(horas)));
            parrafo2.Alignment = Alignment.right;
            parrafo2.Bold();
            parrafo2.FontSize(8);
            parrafo2.Font(new FontFamily("Arial"));

            Novacode.Border border = new Border();
            border.Tcbs = Novacode.BorderStyle.Tcbs_none;
            tabla.SetBorder(TableBorderType.Right, border);
            tabla.SetBorder(TableBorderType.Left, border);
            tabla.SetBorder(TableBorderType.Top, border);
            tabla.SetBorder(TableBorderType.Bottom, border);
            tabla.SetBorder(TableBorderType.InsideH, border);
            tabla.SetBorder(TableBorderType.InsideV, border);

            tabla.SetWidths(new float[]{
                ANCHO*500,
                ANCHO*500
            });
            Documento.InsertTable(tabla);
        }

        public void InsertarHoras(double horas)
        {
            Formatting formato = new Formatting()
            {
                Bold = true,
                FontFamily = new FontFamily("Arial"),
                Size = 8
            };
        }

        public void InsertarTextoReporte(string texto)
        {
            Formatting formato = new Formatting()
            {
                FontFamily = new FontFamily("Trebuchet MS"),
                Size = 10
            };

            InsertarEntrada(texto, formato);
        }

        public void FindAndReplace(string findText, string replaceText, Formatting formato = null)
        {
            try
            {
                this.findAndReplace(Documento, findText, replaceText, formato);
            }
            catch (Exception ex)
            {
                throw new System.ArgumentException("Error generado al intentar reemplazar texto. Descripción: " + ex.Message);
            }
        }

        private void findAndReplace(DocX wordApp, string findText, string replaceText, Formatting formato = null)
        {
            try
            {
                if (formato == null)
                    wordApp.ReplaceText(findText, replaceText);
                else
                    wordApp.ReplaceText(findText, replaceText, false, System.Text.RegularExpressions.RegexOptions.None, formato);
            }
            catch (Exception ex)
            {
                throw new System.ArgumentException(ex.Message);
            }
        }

        public void Dispose()
        {
            Documento.Dispose();
        }
    }

}

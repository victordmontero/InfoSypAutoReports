using LibGit2Sharp;
using System;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using System.Threading;
using System.Net.Mail;
using System.Net;
using System.Security;

namespace InfosypReport
{
    class Program
    {
        private static readonly string OUTPUTDIR = ConfigurationManager.AppSettings["Reportes"];
        private static readonly string REPOSITORIO = ConfigurationManager.AppSettings["Repositorio"];
        private static readonly string NOMBREARCHIVO = ConfigurationManager.AppSettings["NombreArchivo"];
        private static readonly string PLANTILLA = ConfigurationManager.AppSettings["Plantilla"];

        static void Main(string[] args)
        {
            args = (args.Count() == 0) ? new string[1] : args;
            using (var p = new ProcesadorWord(args.Count() < 2 ? OUTPUTDIR + PLANTILLA : args[0]))
            {
                var parrafos = p.Documento.Paragraphs.ToList();

                var commits = ObtenerCommits(args.Count() < 2 ? REPOSITORIO : args[1])
                    .Where(x => x.Key >= DateTime.Today.Subtract(TimeSpan.FromDays(4)) && x.Key <= DateTime.Today);

                for (int i = 0; i < commits.Count(); i++)
                {
                    if (i >= 5) break;
                    var kp = commits.ElementAt(i);
                    p.InsertarFechaCabezera(kp.Key, 8.00);
                    p.InsertarTextoReporte((kp.Value.Contains("no message") ||
                        kp.Value.Contains("Merge branch") ||
                        kp.Value.Count() < 1) ? GenerarMensajePorDefecto() : kp.Value);
                }

                Rellenar(p, commits, DateTime.Today.Subtract(TimeSpan.FromDays(4)));

                ReemplazarFechas(p, desde: DateTime.Today.Subtract(TimeSpan.FromDays(4)), hasta: DateTime.Today);
                p.FindAndReplace("{Nombre}", ConfigurationManager.AppSettings["Nombre"]);
                p.FindAndReplace("{Proyecto}", ConfigurationManager.AppSettings["Proyecto"]);

                p.GuardarComo(string.Format(OUTPUTDIR + NOMBREARCHIVO, DateTime.Today.ToString("dd-MM-yyyy")));
                var fStream = ObtenerDocumentoReciente(OUTPUTDIR);
                Console.WriteLine(EnviarCorreo(ConfigurationManager.AppSettings["Destinatario"], "Reporte", "", fStream));
            }
            Console.Read();
        }

        private static void Rellenar(ProcesadorWord p, IEnumerable<KeyValuePair<DateTime, string>> commits, DateTime fecha)
        {
            if (commits.Count() < 5)
            {
                for (int i = 0; i < 5 - commits.Count(); i++)
                {
                    p.InsertarFechaCabezera(fecha, 8.00);
                    p.InsertarTextoReporte(GenerarMensajePorDefecto());
                    fecha = fecha.AddDays(1);
                }
            }
        }

        private static void ReemplazarFechas(ProcesadorWord p, DateTime desde, DateTime hasta)
        {
            string formato = "dd/MM/yyyy";
            p.FindAndReplace("{Desde}", desde.ToString(formato));
            p.FindAndReplace("{Hasta}", hasta.ToString(formato));
        }

        private static FileStream ObtenerDocumentoReciente(string dir)
        {
            string archivoReciente = string.Empty;

            try
            {
                archivoReciente = Directory.EnumerateFiles(dir)
                        .Select(f => new FileInfo(f))
                        .Where(a => !a.Name.Contains(PLANTILLA))
                        .OrderBy(c => c.CreationTime)
                        .Select(x => x.Name).First();

                return File.Open(dir + archivoReciente, FileMode.Open);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al buscar archivo reciente: " + ex.Message);
                return null;
            }
        }

        private static bool EnviarCorreo(string p1, string p2, string p3, params FileStream[] fStream)
        {
            string host = ConfigurationManager.AppSettings["Host"];
            string pass = ConfigurationManager.AppSettings["Pass"];
            string servidor = ConfigurationManager.AppSettings["Servidor"];
            string dominio = ConfigurationManager.AppSettings["Dominio"];
            int puerto = int.Parse(ConfigurationManager.AppSettings["Puerto"]);

            try
            {
                return EnviarCorreo(host, p1, p2, p3, host, pass, servidor, puerto, fStream);
            }
            catch (Exception ex)
            {
                Console.WriteLine("No se pudo enviar el correo: " + ex.Message);
                return false;
            }
        }

        private static bool EnviarCorreo(string remitente, string destinatario, string titulo, string cuerpo, string usr, string pass, string smtp, int puerto, params FileStream[] adjuntos)
        {
            SmtpClient client = null;
            MailMessage mensaje = null;
            try
            {
                mensaje = new MailMessage(remitente, destinatario, titulo, cuerpo);
                foreach (var adjunto in adjuntos)
                {
                    mensaje.Attachments.Add(new Attachment(adjunto, adjunto.Name.Split('\\').Last()));
                }
                client = new SmtpClient(smtp, puerto);
                client.Credentials = new NetworkCredential(usr, pass);
                client.EnableSsl = true;
                client.Send(mensaje);
                return true;
            }
            catch (Exception)
            {
                throw;
                return false;
            }
        }

        private static string GenerarMensajePorDefecto()
        {
            string[] mensajes = new string[] {
                ConfigurationManager.AppSettings["MensajePorDefecto1"],
                ConfigurationManager.AppSettings["MensajePorDefecto2"],
                ConfigurationManager.AppSettings["MensajePorDefecto3"],
                ConfigurationManager.AppSettings["MensajePorDefecto4"]
            };
            Random rnd = new Random();
            return mensajes[rnd.Next(3)];
        }

        static Dictionary<DateTime, string> ObtenerCommits(string repoDir)
        {
            try
            {
                Dictionary<DateTime, string> commits = new Dictionary<DateTime, string>();
                using (var R = new Repository(repoDir))
                {
                    foreach (var c in R.Commits)
                    {
                        commits.Add(c.Committer.When.DateTime, c.Message);
                    }
                }
                return commits;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Hubo un error al obtener commits: " + ex.Message);
                return new Dictionary<DateTime, string>();
            }
        }

    }
}

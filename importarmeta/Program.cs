using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace importarmeta
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            try { 
            string usuariowinthor   = args[0];
            string usuariobanco     = args[1];
            string banco            = args[2];
            string senhabanco       = args[3];
            string numerorotina     = args[4];


            string sourceDirectory = @"P:\\PCCFM\\PCCFM9806";
            string destinationDirectory = @"C:\\WinThor\\PROD\\PCCFM";

            try
            {
                CopyDirectory(sourceDirectory, destinationDirectory);
                Console.WriteLine("Todos os arquivos foram copiados com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao copiar os arquivos: " + ex.Message);
            }

                Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(usuariowinthor, usuariobanco, banco, senhabanco, numerorotina));
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);    
            }


        }

        static void CopyDirectory(string sourceDir, string destDir)
        {
            // Certifique-se de que o diretório de destino existe
            if (!Directory.Exists(destDir))
            {
                Directory.CreateDirectory(destDir);
            }

            // Copia todos os arquivos do diretório atual (incluindo arquivos no root)
            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(file);
                string destFile = Path.Combine(destDir, fileName);

                // Copia o arquivo para o destino, sobrescrevendo se necessário
                File.Copy(file, destFile, true);
            }

            // Recursivamente copia os subdiretórios
            foreach (string subdir in Directory.GetDirectories(sourceDir))
            {
                string subdirName = Path.GetFileName(subdir);
                string destSubDir = Path.Combine(destDir, subdirName);

                CopyDirectory(subdir, destSubDir);
            }
        }

    }
}

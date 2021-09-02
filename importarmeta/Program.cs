using System;
using System.Collections.Generic;
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
            arquivos("ExcelDataReader.DataSet.dll");
            arquivos("ExcelDataReader.dll");
            arquivos("Oracle.DataAccess.dll");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(usuariowinthor, usuariobanco, banco, senhabanco, numerorotina));
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);    
            }


        }

        private static void arquivos( string nomearquivo)
        {/*
            if (System.IO.File.Exists(@"C:\WinThor\Prod\GLMEP\" + nomearquivo) == true)
            {
             //   MessageBox.Show("j");
            }
            else
            {
                System.IO.File.Copy(System.IO.Path.Combine(@"P:\GLMEP", nomearquivo), System.IO.Path.Combine(@"C:\WinThor\Prod\GLMEP", nomearquivo), true);

            }*/
        }


    }
}

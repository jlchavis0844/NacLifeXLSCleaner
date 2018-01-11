using System;
using System.Windows.Forms;

namespace NacLifeXLSCleaner {
    class Program {
        static string filePath = "";
        static string fileName = "";

        [STAThread]
        static void Main(string[] args) {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = "P:\\RALFG\\Common Files\\Commissions & Insurance\\Commission Statements\\" +
                DateTime.Now.Year.ToString() + "\\";
            ofd.Filter = "XLS files|*.XLS";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK) {
                filePath = ofd.FileName;
                fileName = System.IO.Path.GetFileName(filePath);
                new NacLifeXLS(filePath);
            }
        }
    }
}

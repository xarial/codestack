using System;
using System.Linq;
using System.Windows.Forms;

namespace SwPdmVaultBrowser
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var path = "";

            if (args.Any())
            {
                path = args.First();
            }

            var dlg = new OpenFileDialog()
            {
                InitialDirectory = path,
                Multiselect = true 
            };

            dlg.ShowDialog();
        }
    }
}

using EPDM.Interop.epdm;
using System;

namespace CodeStack.ListPdmVaults
{
    class Program
    {
        static void Main(string[] args)
        {
            var vault = new EdmVault5Class();
            EdmViewInfo[] views;
            vault.GetVaultViews(out views, false);

            foreach (var view in views)
            {
                Console.WriteLine($"{view.mbsVaultName}:{view.mbsPath}");
            }
        }
    }
}

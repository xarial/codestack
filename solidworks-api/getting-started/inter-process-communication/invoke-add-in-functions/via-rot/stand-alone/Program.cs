using CodeStack.GeometryHelper;
using System;
using System.Diagnostics;
using System.Linq;

namespace StandAlone
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                var minArea = double.Parse(args[0]);

                var swPrcId = Process.GetProcessesByName("SLDWORKS").First().Id;

                var geomHelperFactory = new GeometryHelperApiObjectFactory();

                var geomHelperApi = geomHelperFactory.GetInstance(swPrcId);

                var count = geomHelperApi.GetFacesCount(minArea);

                Console.WriteLine($"Selected body contains {count} faces of area more or equal to {minArea}");
            }
            catch(Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write(ex.Message);
                Console.ResetColor();
            }
        }
    }
}

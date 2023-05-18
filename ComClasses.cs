using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace NewComServer
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary
    {
        string ComLibraryHello();
        double Add(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary))]
    public class ComLibrary : IComLibrary
    {
        public string ComLibraryHello()
        {
            return "Hello from DnaComServer.ComLibrary at " + ExcelDnaUtil.XllPath ;
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
    }

    public class ComLibrary2
    {
        public string ComLibrary2Hello()
        {
            return "Hello from DnaComServer.ComLibrary2";
        }

        public double Add2(double x, double y)
        {
            return x + y;
        }
    }
}

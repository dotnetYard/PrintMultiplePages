using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintMultiplePages
{
    class Program
    {
        static void Main(string[] args)
        {
            PrintProcecssor procecssor = new PrintProcecssor();
            procecssor.PrintReport(@"PRINTER NAME", // Enter your printer name
                "REPORT NAME WITH PATH", // Enter your report name with path
                1,                
                "REPORT PARAMETER", // Enter your report parameter
                false, "Force 8.5 X 11");
        }
    }
}

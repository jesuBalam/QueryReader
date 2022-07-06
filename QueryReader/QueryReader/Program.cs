using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QueryReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Press any key to start.");
            Console.ReadKey();

            QueryReader.ReadFile(@ConfigurationManager.AppSettings["FileInput"]);
            Console.WriteLine("Process Done.");
            Console.ReadKey();
        }
    }
}

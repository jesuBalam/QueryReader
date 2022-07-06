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

            Console.WriteLine("Paste file path to read: (Space/Enter to take default path in config)");
            string pathFile = Console.ReadLine();
            pathFile.Trim();

            if(!string.IsNullOrEmpty(pathFile))
            {
                ConfigurationManager.AppSettings["FileInput"] = pathFile;
            }

            QueryReader.ReadFile(@ConfigurationManager.AppSettings["FileInput"]);
            Console.WriteLine("Process Done.");
            Console.ReadKey();
        }
    }
}

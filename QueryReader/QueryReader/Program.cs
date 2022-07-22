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
            try
            {
                Console.WriteLine("Press any key to start.");
                Console.ReadKey();

                #region Execute querys from file
                Console.WriteLine("Paste file path to read: (Space/Enter to take default path in config)");
                string pathFile = Console.ReadLine();
                pathFile.Trim();

                if (!string.IsNullOrEmpty(pathFile))
                {
                    ConfigurationManager.AppSettings["FileInput"] = pathFile;
                }
                QueryReader.ReadFile(@ConfigurationManager.AppSettings["FileInput"]);
                #endregion                
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }            
        }
    }
}

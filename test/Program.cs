using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using EEAuto;

namespace test
{
    class Program
    {
        static void Main(string[] args)
        {
            StreamReader sr = new StreamReader(@"C:\EE\1.nxl.txt");
            string line = "";
            while ((line = sr.ReadLine()) != null)
            {
                if (line.ToLower().Contains(" "))
                {
                    Console.WriteLine("hello");
                }
            }
                  

            
            sr.Close();
            Console.ReadKey();
        }
    }
}

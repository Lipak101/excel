using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO.Compression;
using System.Text.RegularExpressions;


namespace _21_11_23
{
    class Program
    {
        static void Main(string[] args)
        {

            ////Console.WriteLine("Ez itt a meno beno titkositast feloldo programunk !");
            ////Console.WriteLine("Írja be a file nevét:");
            ////string file_name = Console.ReadLine();
            ////string file = @"" + file_name + ".xlsx";
            ////Átalakit(file_name);
            ////Console.WriteLine("Átalakítás megtörtént!");
            ////Console.ReadKey();
            ////Kitomorit(file_name);
            ////Console.WriteLine("Tömörítés kibontása megtörtént!");
            ////Console.ReadKey();
            ////Kiolvas("vedett2/xl/worksheets/sheet1.xml");
            ////Console.WriteLine("Titkosítás feltörése megtörtént!");
            ////Console.ReadKey();
            Átalakit2("vedett2");
            Console.WriteLine("Átalakítás megtörtént!");
            Console.ReadKey();



        }
        static void Átalakit2(string file_name)
        {
            string input = file_name;
            File.Move(input, Path.ChangeExtension(input, ".zip"));

        }
        static void Átalakit(string file_name)
        {
            string input = @"" + file_name + ".xlsx";
            File.Move(input, Path.ChangeExtension(input, ".zip"));

        }
        static void Kitomorit(string file_name)
        {
            string zipPath = @"" + file_name + ".zip";
            string extractPath = @"C:\Users\71617410798\source\repos\ConsoleApp2\ConsoleApp2\bin\Debug\net5.0";
            ZipFile.ExtractToDirectory(zipPath, extractPath);
        }
        static void Kiolvas(string fajl)
        {
            string szoveg = File.ReadAllText(fajl);
            string[] szavak = szoveg.Split('<');
            for (int j = 0; j < szavak.Length; j++)
            {
                if (szavak[j].Contains("sheetProtection") || szavak[j].Contains("pageMargins"))
                {
                    szavak[j] = "";
                }
            }
            szoveg = "";
            for (int i = 0; i < szavak.Length; i++)
            {
                if (szavak[i] != "")
                {
                    szoveg += "<" + szavak[i];
                }
            }
            File.WriteAllText(fajl, szoveg);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace sigscan
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                System.Console.WriteLine("Por favor escriba un argumento");
                System.Console.WriteLine("Uso: sigscan <path>");
                return 1;
            }

            string path = string.Join("", args);

            string[] files = Directory.GetFiles(@path, "*.dll");
            string[] listA = new string[files.Length];
            string[] listB = new string[files.Length];
            for (int i = 0; i < files.Length; i++)
            {
                FileVersionInfo fileVer = FileVersionInfo.GetVersionInfo((files[i]));
                Console.WriteLine(fileVer.FileDescription + "\t" + fileVer.FileVersion);
                listA[i] = fileVer.FileDescription;
                listB[i] = fileVer.FileVersion;
            }

            //Exportar a Excel
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return 1;
            }
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            // Selects excel cells within the range of the list of files
            Range aRange = ws.get_Range("A1", "A" + files.Length);
            Range bRange = ws.get_Range("B1", "B" + files.Length);

            //aRange.Cells[1, "A"].Value2 = "NOMBRE";
            //bRange.Cells[1, "B"].Value2 = "VERSION";

            if (aRange == null)
            {
                Console.WriteLine("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }

            for (int k = 1; k < files.Length + 1; k++)
            {
                aRange.Cells[k, "A"].Value2 = listA[k - 1];
                bRange.Cells[k, "B"].Value2 = listB[k - 1];
            }

            return 0;
        }
    }
}
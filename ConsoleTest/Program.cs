using OpenXMLImportDLL;
using System;
using System.IO;

/*using System;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using System.Collections;
using System.Linq;*/

namespace ConsoleApplication1
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Enter PATH");
            string path = "D:/hTestFormula";
            AddDir(path);
            AddExcel(path);
        }
        static string AddDir(string path)
        {

            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    Console.WriteLine("That path exists already.");
                    return (path);
                }

                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
                return (path);
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
                return ("Error");
            }
            finally { }

        }

        static void AddExcel(string path)
        {

            for (int k = 1; k <2; k++)
            {
            for (int i = 1; i < 8; i++)
            {
                for (int j = 1; j < 80; j++)
                {
                    OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i.ToString() + j.ToString(), 6);

                }
            }

            OpenXMLImportDLL.ExcelImport.AddCellData(8, 1, "=A1+A2+A3+A4+A5+A6+A7*A1", 1);
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 2, "=B1+B2+B3+B4+B5+B6+B7*A2", 2);
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 3, "=C1+C2+C3+C4+C5+C6+C7", 3);
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 4, "=D1+D2+D3+D4+D5+D6+D7", 4);
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 5, "=E1+E2+E3+E4+E5+E6+E7", 5);
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 6, "=F1+F2+F3+F4+F5+F6+F7", 6);
     
            OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 50);
            OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 30);

                OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest.xlsx");
                OpenXMLImportDLL.ExcelImport.ClearArray();
            }
        }
    }

}
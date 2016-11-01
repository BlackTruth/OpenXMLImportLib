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
            string path = "D:/TestFormula";
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
            //for (int i = 1; i < 8; i++)
            //{
            //    for (int j = 1; j < 8; j++)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, "QWdasd111");

            //    }
            //}
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "1");
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, "2");
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 3, "3");
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 4, "4");
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 1, "1,1");
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 2, "2,2");
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 3, "3,3");
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 4, "4,4");
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "=A1+A2");
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "=B1+B2");
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 3, "=C1+C2");
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "=D1+D2");

            OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest.xlsx");
            OpenXMLImportDLL.ExcelImport.ClearArray();
        }
    }

}

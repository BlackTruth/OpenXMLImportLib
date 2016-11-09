using OpenXMLImportDLL;
using System;
using System.IO;
using System.Text.RegularExpressions;

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
            string path = "D:/Test";
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
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, "0", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "1", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(5, 3, "2", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(7, 4, "3", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(9, 5, "4", 5);
            OpenXMLImportDLL.ExcelImport.AddCellData(11, 6, "5", 5);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "АЛИГМЕЕЕНТ", 5);
            OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 40);

            OpenXMLImportDLL.ExcelImport.AddMergeCell("D2:E2");

            OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 40);
            OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest.xlsx");
            OpenXMLImportDLL.ExcelImport.ClearArray();

        }
    }

}
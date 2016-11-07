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
            string path = "E:/Test";
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
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 1,null, 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, "", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 3, "12", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 4, "", 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 5, "1", 1);
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 1, null, 6);
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 2, "4", 3);
         //   OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 40);
                OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest.xlsx");
                OpenXMLImportDLL.ExcelImport.ClearArray();
            
        }
    }

}
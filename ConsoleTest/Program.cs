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
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "Test!!!", true, 21, false, 5, "Calibri", false, true, 0, 0, false);
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 1, "241241.12", false, 11, false, 1, "Tahoma", false, false, 1, 1, true);
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "13433,14", false, 13, true, 2, "Tahoma", true, true,2, 2, true);
            OpenXMLImportDLL.ExcelImport.AddCellData(4, 1, "4", true, 12, false, 3, "Tahoma", true, true, 3, 3, false);
            OpenXMLImportDLL.ExcelImport.AddCellData(5, 1, "5", true, 11, true, 4, "Times New Roman", true, false, 3,2, false);
            OpenXMLImportDLL.ExcelImport.AddCellData(6, 1, "6", true, 14, false, 5, "Times New Roman", true, true, 3, 1, false);
           OpenXMLImportDLL.ExcelImport.AddCellData(7, 1, "7", false, 12, true, 5, "Comic Sans MS", true, false, 2, 1, true);
            OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest.xlsx");
            
        }
    }

}
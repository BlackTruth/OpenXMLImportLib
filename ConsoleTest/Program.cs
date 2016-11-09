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
            OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 20);
            OpenXMLImportDLL.ExcelImport.AddRowHeight(2, 20);
            OpenXMLImportDLL.ExcelImport.AddRowHeight(3, 20);
            for (int q = 1; q <= 1; q++)
            {


                for (int i = 1; i < 7; i++)
                {
                    for (int j = 1; j < 2; j++)
                    {
                        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, (i + j).ToString(), 5);
                    }
                }
                OpenXMLImportDLL.ExcelImport.AddCellData(7, 2, "=A1+A2", 6);
                OpenXMLImportDLL.ExcelImport.AddMergeCell("A1:A2");
                OpenXMLImportDLL.ExcelImport.AddCellData(7, 1, "=SUM(R[-6]C:R[-1]C)", 6);
                OpenXMLImportDLL.ExcelImport.AddMergeCell("C1:C2");
               OpenXMLImportDLL.ExcelImport.AddMergeCell("E1:K12");
               OpenXMLImportDLL.ExcelImport.AddColumnWidth(5, 100);
                OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest"+q+".xlsx");
                OpenXMLImportDLL.ExcelImport.ClearArray();
            }

        }
    }

}
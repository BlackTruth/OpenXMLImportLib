﻿using OpenXMLImportDLL;
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
            string path = "C:/Test";
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
            //for (int i = 1; i <= 100; i++)
            //    for (int j = 1; j <= 100; j++)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i + "|" + j, false, 10, false, 5, "Viner Hand ITC", false, true, 1, 1, false);
            //    }

            //for (int i = 110; i >= 1; i--)
            //    for (int j = 110; j >= 1; j--)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i + "|" + j, false, 15, false, 5, "Old English Text MT", false, true, 1, 1, false);
            //    }

            //for (int i = 200; i >= 150; i--)
            //    for (int j = 300; j >= 150; j--)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, "Row" + i + "|" + j + "Column", false, 12, true, 5, "Old English Text MT", false, true, 1, 1, false);
            //    }

            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);

            //OpenXMLImportDLL.ExcelImport.AddMergeCell("A5:F6"); OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 11, false, 4, "Tahoma", false, false, 0, 0, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 11, false, 4, "Tahoma", false, false, 0, 0, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 11, false, 4, "Tahoma", false, false, 0, 0, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", true, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 15, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 11, false, 4, "Tahoma", false, false, 0, 0, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(4, 4, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 2, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 11, false, 4, "Tahoma", false, false, 0, 0, true);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 2, null, false, 11, false, 5, "Tahoma", false, false, 4, 4, false);
            //OpenXMLImportDLL.ExcelImport.AddCellData(5, 5, "=A1+A2+3", true, 11, true, 5, "Tahoma", true, true, 1, 1, true);
            OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "=sum(R5C)", false, 11, false, 4, "Tahoma", false, false, 2, 2, false, "FF1B1B");
            OpenXMLImportDLL.ExcelImport.AddCellData(2, 1, "3", false, 11, false, 4, "Tahoma", false, false, 3, 3, false, "FF1B1B");
            OpenXMLImportDLL.ExcelImport.AddCellData(3, 1, "111213,1010", false, 15, false, 4, "Tahoma", false, false, 0, 0, true, "FF1B1B");
            OpenXMLImportDLL.ExcelImport.AddCellData(4, 1, "3pipo", false, 11, false, 4, "Tahoma", false, false, 4, 4, false, "FF1B1B");
            OpenXMLImportDLL.ExcelImport.AddCellData(5, 1, "", false, 11, false, 5, "Tahoma", false, false, 4, 4, false, "FF1B1B");
            OpenXMLImportDLL.ExcelImport.AddCellData(6, 1, null, false, 15, false, 3, "Times New Roman", false, false, 3, 3, false, "FFFFFF");
            OpenXMLImportDLL.ExcelImport.AddCellData(7, 1, "3", false, 11, false, 4, "Tahoma", false, false, 0, 0, false, "FFFFFF");
            OpenXMLImportDLL.ExcelImport.AddCellData(8, 1, "111213,1010", false, 15, false, 4, "Tahoma", false, false, 1, 1, true, "FFFFFF");
            OpenXMLImportDLL.ExcelImport.AddCellData(9, 1, "3pipo", false, 11, false, 4, "Times New Roman", false, false, 3, 3, false, "FFFFFF");
            OpenXMLImportDLL.ExcelImport.AddCellData(10, 1, "0000085516", false, 110, false, 110, "Tahoma", false, false, 200, 200, false, "FFFFFF");


            OpenXMLImportDLL.ExcelImport.AddCellData(11, 1, "/T00505050", false, 23, true, 4, "Tahoma", false, false, 4, 4, true, "1BFF60");


            OpenXMLImportDLL.ExcelImport.AddCellData(12, 1, "110", false, 23, true, 4, "Tahoma", false, false, 4, 4, true, "FFFFFF");
                     //OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 30);
            //OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 30);
            //OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 40);
            //OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 40);
            
           

            OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/A1.xlsx");


            //int q=OpenXMLImportDLL.ExcelImport.GetColRow();
            //Console.WriteLine(q);
            //Console.ReadKey();
        }
    }

}
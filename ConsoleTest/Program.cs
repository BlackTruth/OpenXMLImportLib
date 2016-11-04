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
            ExcelImport obj = ExcelImport.getInstance();
           
            for (int i = 1; i < 100; i++)
            {
                for (int j = 1; j < 10; j++)
                {
                    obj.AddCellData(i, j, i + "|" + j, 6);
                }
            }
            obj.AddCellData(2, 2, "2", 2);
            obj.AddColumnWidth(1, 100);
            obj.AddRowHeight(1, 100);

            obj.GenerateExcel(path + "/Newtest2.xlsx");
            obj.ClearArray();
            //for (int i = 1; i < 100; i++)
            //{
            //    for (int j = 1; j < 10; j++)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i + "|" + j, 6);
            //    }
            //}
            //OpenXMLImportDLL.ExcelImport.AddColumnWidth(3, 100);
            //OpenXMLImportDLL.ExcelImport.AddRowHeight(3, 200);
            //OpenXMLImportDLL.ExcelImport.AddRowHeight(7, 200);
            //OpenXMLImportDLL.ExcelImport.AddCellData(2, 2, "1", 6);
            //OpenXMLImportDLL.ExcelImport.AddCellData(3, 4, "2", 6);
            //OpenXMLImportDLL.ExcelImport.AddCellData(2, 1, "3", 6);
            //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "4", 6);
            //OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/Newtest2.xlsx");
            //OpenXMLImportDLL.ExcelImport.ClearArray();
        }
    }

}
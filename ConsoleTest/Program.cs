using System;
using System.IO;
using System.Threading;
using System.Drawing;
using System.Windows.Media;
using System.Windows.Media.Imaging;



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

        private static Bitmap BitmapImage2Bitmap(BitmapSource bitmapImage)
        {
            using (var outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapImage));
                enc.Save(outStream);
                var bitmap = new Bitmap(outStream);
                return bitmap;
            }
        }


        static void AddExcel(string path)
        {
            //OpenXMLImportDLL.ExcelImport.AddPageSetup(true, 1, 1, false);

            //for (int i = 1; i <=20; i++)
            //{
            //    for (int j = 1; j <= 20; j++)
            //    {
            //        var url = new Uri("C:\\test\\1.png");
            //        var bi = new BitmapImage(url);
            //        var bm = BitmapImage2Bitmap(bi);
            //        var pixel = bm.GetPixel(i, j);
            //        string hex = pixel.R.ToString("X2") + pixel.G.ToString("X2") + pixel.B.ToString("X2");
            //        OpenXMLImportDLL.ExcelImport.AddCellData(j, i, "", false, 1, false, 0, "Viner Hand ITC", false, true, 1, 1, false, hex);
            //    }
            //}

            //for (int i = 1; i <= 120; i++)
            //{
            //    for (int j = 1; j <= 120; j++)
            //    {
            //        var url = new Uri("C:\\test\\2.png");
            //        var bi = new BitmapImage(url);
            //        var bm = BitmapImage2Bitmap(bi);
            //        var pixel = bm.GetPixel(i, j);
            //        string hex = pixel.R.ToString("X2") + pixel.G.ToString("X2") + pixel.B.ToString("X2");
            //        OpenXMLImportDLL.ExcelImport.AddCellData(j, i, "", false, 1, false, 0, "Viner Hand ITC", false, true, 1, 1, false, hex);
            //    }
            //}

            //for (int i = 1; i <= 20; i++)
            //{
            //    for (int j = 1; j <= 20; j++)
            //    {
            //        var url = new Uri("C:\\test\\1.png");
            //        var bi = new BitmapImage(url);
            //        var bm = BitmapImage2Bitmap(bi);
            //        var pixel = bm.GetPixel(i, j);
            //        string hex = pixel.R.ToString("X2") + pixel.G.ToString("X2") + pixel.B.ToString("X2");
            //        OpenXMLImportDLL.ExcelImport.AddCellData(j, i, "", false, 1, false, 0, "Viner Hand ITC", false, true, 1, 1, false, hex);
            //    }
            //}


            //for (int i = 1; i <=120; i++)
            //{


            //    for (int j = 1; j <=120; j++)
            //    {
            //        OpenXMLImportDLL.ExcelImport.AddColumnWidth(j, 1);
            //        OpenXMLImportDLL.ExcelImport.AddRowHeight(i, 6);
            //    }
            //}


            for (int p = 1; p <= 1; p++)
            {
                //Random rnd = new Random();
                ////OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "№", false, 10, false, 5, "Viner Hand ITC", false, true, 1, 1, false, "FFFFFF");
                ////OpenXMLImportDLL.ExcelImport.AddMergeCell("F5:F6");
                ////OpenXMLImportDLL.ExcelImport.AddMergeCell("a1:E4");
                ////OpenXMLImportDLL.ExcelImport.AddMergeCell("F11:P16");
                //for (int i = 1; i <= 150000; i++)
                //    for (int j = 1; j <= 10; j++)
                //    {
                //        //int row = rnd.Next(1, 1000);
                //        //int column = rnd.Next(1, 1000);
                //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i + "|" + j, false, 10, false, 5, "Viner Hand ITC", false, true, 1, 1, false, "FFFFFF");
                //    }

                //for (int i = 110; i >= 1; i--)
                //    for (int j = 110; j >= 1; j--)
                //    {
                //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, i + "|" + j, false, 15, false, 5, "Old English Text MT", false, true, 1, 1, false, "FFFFFF");
                //    }

                //for (int i = 200; i >= 150; i--)
                //    for (int j = 300; j >= 150; j--)
                //    {
                //        OpenXMLImportDLL.ExcelImport.AddCellData(i, j, "Row" + i + "|" + j + "Column", false, 12, true, 5, "Old English Text MT", false, true, 1, 1, false, "FFFFFF");
                //    }


                //OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 30);
                //OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 30);
                //OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 40);
                //OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 40);
                //OpenXMLImportDLL.ExcelImport.AddMergeCell("F5:F6");

                //OpenXMLImportDLL.ExcelImport.AddCellData(1, 1, "/T123,1942", false, 10, false, 5, "Viner Hand ITC", false, true, 3,3, false, "FFFFFF");
                OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/A" + p + ".xlsx");


                //int q=OpenXMLImportDLL.ExcelImport.GetColRow();
                //Console.WriteLine(q);
                //Console.ReadKey();
            }
        }

    }
}
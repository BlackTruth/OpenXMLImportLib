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

        static string Colorrrr()
        {
            var random = new Random();
            string color = (String.Format("{0:X6}", random.Next(0x1000000))).ToString();
            string s = color.ToString();
            return s;
        }
        static void AddExcel(string path)
        {


            for (int i = 1; i <=127 ; i++)
            {

                OpenXMLImportDLL.ExcelImport.AddRowHeight(i, 5);
                for (int j = 1; j <= 127; j++)
                {
                    string s = Colorrrr();
                    var url = new Uri("C:\\test\\3.png");
                    var bi = new BitmapImage(url);
                    var bm = BitmapImage2Bitmap(bi);
                    var pixel = bm.GetPixel(i, j);
                    string hex = pixel.R.ToString("X2") + pixel.G.ToString("X2") + pixel.B.ToString("X2");
                        OpenXMLImportDLL.ExcelImport.AddCellData(j, i, "", false, 1, false, 5, "Viner Hand ITC", false, true, 1, 1, false, hex);
                        OpenXMLImportDLL.ExcelImport.AddColumnWidth(j, 1);

                }
            }


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
            //OpenXMLImportDLL.ExcelImport.AddRowHeight(1, 40);
            //OpenXMLImportDLL.ExcelImport.AddColumnWidth(1, 40);



            OpenXMLImportDLL.ExcelImport.GenerateExcel(path + "/A1.xlsx");


            //int q=OpenXMLImportDLL.ExcelImport.GetColRow();
            //Console.WriteLine(q);
            //Console.ReadKey();
        }
    }

}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using excel = Microsoft.Office.Interop.Excel;




namespace OpenXMLImportDLL
{
    public class ExcelImport
    {
        static SortedList<int, double?> columnWidthArr = new SortedList<int, double?>();
        static SortedList<int, double?> rowHeightArr = new SortedList<int, double?>();
        static List<string> mergeArr = new List<string>();
        static List<CellData> cellsData = new List<CellData>();
        static List<CellStyleFormat> cellStyleFormatList = new List<CellStyleFormat>();
        static List<FontStyleFormat> fontStyleFormatList = new List<FontStyleFormat>();

        //Convert Excel column number to Excel column name (1=A, 2=B)
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modul;
            while (dividend > 0)
            {
                modul = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modul).ToString() + columnName;
                dividend = (int)((dividend - modul) / 26);
            }
            return columnName;
        }

        private static int GetExcelColumnNumber(string columnName)
        {
            string output = regexReplace(columnName, 1);
            if (string.IsNullOrEmpty(output)) throw new ArgumentNullException("columnName");
            output = output.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < output.Length; i++)
            {
                sum *= 26;
                sum += (output[i] - 'A' + 1);
            }
            return sum;
        }

        private static string regexReplace(string input, int sw)
        {
            //sw=1 remove numbers from string //sw=2 remove char from string
            if (sw == 1)
            {
                string output = Regex.Replace(input, @"[\d-]", string.Empty);
                return output;
            }
            if (sw == 2)
            {
                string output = Regex.Replace(input, @"[A-Z]", string.Empty);
                return output;
            }
            return "error";
        }

        // Adds child parts and generates content of the specified part.

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int GenerateExcel(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                //ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                //GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

                WorkbookPart workbookPart = document.AddWorkbookPart();
                GenerateWorkbookPartContent(workbookPart);

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
                GenerateWorkbookStylesPartContent(workbookStylesPart); //Стили ячеек

                //ThemePart themePart1 = workbookPart.AddNewPart<ThemePart>("rId2");
                //GenerateThemePart1Content(themePart1);

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
                GenerateWorksheetPartContent(worksheetPart);

                SharedStringTablePart sharedStringTablePart1 = workbookPart.AddNewPart<SharedStringTablePart>("rId4");
                GenerateSharedStringTablePart1Content(sharedStringTablePart1);

                SetPackageProperties(document);
            }

            ClearArray();
            return 0;
        }

        //Generates content of extendedFilePropertiesPart1.
        //private static void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        //{
        //    Ap.Properties properties1 = new Ap.Properties();
        //    properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
        //    Ap.Application application1 = new Ap.Application();
        //    application1.Text = "Microsoft Excel";
        //    Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
        //    documentSecurity1.Text = "0";
        //    Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
        //    scaleCrop1.Text = "false";

        //    Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

        //    Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

        //    Vt.Variant variant1 = new Vt.Variant();
        //    Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
        //    vTLPSTR1.Text = "Листы";

        //    variant1.Append(vTLPSTR1);

        //    Vt.Variant variant2 = new Vt.Variant();
        //    Vt.VTInt32 vTInt321 = new Vt.VTInt32();
        //    vTInt321.Text = "1";

        //    variant2.Append(vTInt321);

        //    vTVector1.Append(variant1);
        //    vTVector1.Append(variant2);

        //    headingPairs1.Append(vTVector1);

        //    Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

        //    Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
        //    Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
        //    vTLPSTR2.Text = "Лист1";

        //    vTVector2.Append(vTLPSTR2);

        //    titlesOfParts1.Append(vTVector2);
        //    Ap.Company company1 = new Ap.Company();
        //    company1.Text = "";
        //    Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
        //    linksUpToDate1.Text = "false";
        //    Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
        //    sharedDocument1.Text = "false";
        //    Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
        //    hyperlinksChanged1.Text = "false";
        //    Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
        //    applicationVersion1.Text = "15.0300";

        //    properties1.Append(application1);
        //    properties1.Append(documentSecurity1);
        //    properties1.Append(scaleCrop1);
        //    properties1.Append(headingPairs1);
        //    properties1.Append(titlesOfParts1);
        //    properties1.Append(company1);
        //    properties1.Append(linksUpToDate1);
        //    properties1.Append(sharedDocument1);
        //    properties1.Append(hyperlinksChanged1);
        //    properties1.Append(applicationVersion1);

        //    extendedFilePropertiesPart1.Properties = properties1;
        //}

        //Generates content of workbookPart1.
        private static void GenerateWorkbookPartContent(WorkbookPart workbookPart)
        {
            Workbook workbook = new Workbook() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x15" } };
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            FileVersion fileVersion = new FileVersion() { ApplicationName = "xl", LastEdited = "6", LowestEdited = "4", BuildVersion = "14420" };
            WorkbookProperties workbookProperties = new WorkbookProperties() { FilterPrivacy = true, DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews = new BookViews();
            WorkbookView workbookView = new WorkbookView() { XWindow = 240, YWindow = 105, WindowWidth = (UInt32Value)14805U, WindowHeight = (UInt32Value)8010U };

            bookViews.Append(workbookView);

            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet() { Name = "Лист1", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets.Append(sheet);

            CalculationProperties calculationProperties = new CalculationProperties()
            {
                CalculationId = (UInt32Value)152511U,
                ForceFullCalculation = true,
                FullCalculationOnLoad = true
            };



            workbook.Append(fileVersion);
            workbook.Append(workbookProperties);
            workbook.Append(bookViews);
            workbook.Append(sheets);
            workbook.Append(calculationProperties);

            workbookPart.Workbook = workbook;
        }

        // Generates content of workbookStylesPart1.
        private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            int counterFonts = 0;
            Stylesheet stylesheet1 = new Stylesheet();



            Fills fills1 = new Fills();

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);



            Borders borders = new Borders() { Count = (UInt32Value)6U };

            Border noBorder = new Border(); //id 0
            borders.Append(noBorder);

            Border leftBorder = new Border(); //id 1
            Border rightBorder = new Border(); //id 2
            Border topBorder = new Border(); //id 3
            Border botBorder = new Border(); //id 4
            Border allBorder = new Border(); //id 5

            LeftBorder leftBorder1 = new LeftBorder() { Style = BorderStyleValues.Thin };
            leftBorder.Append(leftBorder1);

            RightBorder rightBorder1 = new RightBorder() { Style = BorderStyleValues.Thin };
            rightBorder.Append(rightBorder1);

            TopBorder topBorder1 = new TopBorder() { Style = BorderStyleValues.Thin };
            topBorder.Append(topBorder1);

            BottomBorder bottomBorder1 = new BottomBorder() { Style = BorderStyleValues.Thin };
            botBorder.Append(bottomBorder1);

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            allBorder.Append(leftBorder2);
            allBorder.Append(rightBorder2);
            allBorder.Append(topBorder2);
            allBorder.Append(bottomBorder2);

            borders.Append(leftBorder);
            borders.Append(rightBorder);
            borders.Append(topBorder);
            borders.Append(botBorder);
            borders.Append(allBorder);




            Fonts fonts = new Fonts();

            Font defFont = new Font();
            FontSize dFontSize = new FontSize() { Val = 11D };
            FontName dFontName = new FontName() { Val = "Calibri" };
            defFont.Append(dFontSize);
            defFont.Append(dFontName);
            fonts.Append(defFont);

            foreach (FontStyleFormat f in fontStyleFormatList)
            {

                Font font = new Font();
                if (f.Bold)
                {
                    Bold bold = new Bold();
                    font.Append(bold);
                }
                if (f.Italic)
                {
                    Italic italic = new Italic();
                    font.Append(italic);
                }
                if (f.Underline)
                {
                    Underline underline = new Underline();
                    font.Append(underline);
                }
                FontSize fontSize = new FontSize() { Val = f.Size };
                font.Append(fontSize);

                FontName fontName = new FontName() { Val = f.FontName };
                font.Append(fontName);

                fonts.Append(font);
                counterFonts++;
            }




            CellStyleFormats cellStyleFormats1 = new CellStyleFormats();
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats = new CellFormats();
            CellFormat dCellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            cellFormats.Append(dCellFormat);

            //if (counterFonts > 0)
            //    for (int i = 1; i <= counterFonts; i++)
            //    {
            //        CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)(UInt32)i, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            //        cellFormats.Append(cellFormat);
            //    }

            foreach (CellStyleFormat d in cellStyleFormatList)
            {
                CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)(UInt32)d.FontIndex, FillId = (UInt32Value)0U, BorderId = (UInt32Value)(UInt32)d.LineStyle, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true };
                cellFormats.Append(cellFormat);
               
                if(d.WrapText)
                {
                  Alignment alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };
                  cellFormat.Append(alignment);
                }

            }



            CellStyles cellStyles1 = new CellStyles();
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats();
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

            stylesheet1.Append(fonts);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart.Stylesheet = stylesheet1;





        }

        // Generates content of themePart1.
        //private static void GenerateThemePart1Content(ThemePart themePart1)
        //{
        //    A.Theme theme1 = new A.Theme() { Name = "Тема Office" };
        //    theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        //    A.ThemeElements themeElements1 = new A.ThemeElements();

        //    A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Стандартная" };

        //    A.Dark1Color dark1Color1 = new A.Dark1Color();
        //    A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

        //    dark1Color1.Append(systemColor1);

        //    A.Light1Color light1Color1 = new A.Light1Color();
        //    A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

        //    light1Color1.Append(systemColor2);

        //    A.Dark2Color dark2Color1 = new A.Dark2Color();
        //    A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

        //    dark2Color1.Append(rgbColorModelHex1);

        //    A.Light2Color light2Color1 = new A.Light2Color();
        //    A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

        //    light2Color1.Append(rgbColorModelHex2);

        //    A.Accent1Color accent1Color1 = new A.Accent1Color();
        //    A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

        //    accent1Color1.Append(rgbColorModelHex3);

        //    A.Accent2Color accent2Color1 = new A.Accent2Color();
        //    A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

        //    accent2Color1.Append(rgbColorModelHex4);

        //    A.Accent3Color accent3Color1 = new A.Accent3Color();
        //    A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

        //    accent3Color1.Append(rgbColorModelHex5);

        //    A.Accent4Color accent4Color1 = new A.Accent4Color();
        //    A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

        //    accent4Color1.Append(rgbColorModelHex6);

        //    A.Accent5Color accent5Color1 = new A.Accent5Color();
        //    A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

        //    accent5Color1.Append(rgbColorModelHex7);

        //    A.Accent6Color accent6Color1 = new A.Accent6Color();
        //    A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

        //    accent6Color1.Append(rgbColorModelHex8);

        //    A.Hyperlink hyperlink1 = new A.Hyperlink();
        //    A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

        //    hyperlink1.Append(rgbColorModelHex9);

        //    A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
        //    A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

        //    followedHyperlinkColor1.Append(rgbColorModelHex10);

        //    colorScheme1.Append(dark1Color1);
        //    colorScheme1.Append(light1Color1);
        //    colorScheme1.Append(dark2Color1);
        //    colorScheme1.Append(light2Color1);
        //    colorScheme1.Append(accent1Color1);
        //    colorScheme1.Append(accent2Color1);
        //    colorScheme1.Append(accent3Color1);
        //    colorScheme1.Append(accent4Color1);
        //    colorScheme1.Append(accent5Color1);
        //    colorScheme1.Append(accent6Color1);
        //    colorScheme1.Append(hyperlink1);
        //    colorScheme1.Append(followedHyperlinkColor1);

        //    A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Стандартная" };

        //    A.MajorFont majorFont1 = new A.MajorFont();
        //    A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria", Panose = "020F0302020204030204" };
        //    A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
        //    A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
        //    A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
        //    A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
        //    A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
        //    A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
        //    A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
        //    A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
        //    A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
        //    A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
        //    A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
        //    A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
        //    A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
        //    A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
        //    A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
        //    A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
        //    A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
        //    A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
        //    A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
        //    A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
        //    A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
        //    A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
        //    A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
        //    A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
        //    A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
        //    A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
        //    A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
        //    A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
        //    A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
        //    A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
        //    A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
        //    A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

        //    majorFont1.Append(latinFont1);
        //    majorFont1.Append(eastAsianFont1);
        //    majorFont1.Append(complexScriptFont1);
        //    majorFont1.Append(supplementalFont1);
        //    majorFont1.Append(supplementalFont2);
        //    majorFont1.Append(supplementalFont3);
        //    majorFont1.Append(supplementalFont4);
        //    majorFont1.Append(supplementalFont5);
        //    majorFont1.Append(supplementalFont6);
        //    majorFont1.Append(supplementalFont7);
        //    majorFont1.Append(supplementalFont8);
        //    majorFont1.Append(supplementalFont9);
        //    majorFont1.Append(supplementalFont10);
        //    majorFont1.Append(supplementalFont11);
        //    majorFont1.Append(supplementalFont12);
        //    majorFont1.Append(supplementalFont13);
        //    majorFont1.Append(supplementalFont14);
        //    majorFont1.Append(supplementalFont15);
        //    majorFont1.Append(supplementalFont16);
        //    majorFont1.Append(supplementalFont17);
        //    majorFont1.Append(supplementalFont18);
        //    majorFont1.Append(supplementalFont19);
        //    majorFont1.Append(supplementalFont20);
        //    majorFont1.Append(supplementalFont21);
        //    majorFont1.Append(supplementalFont22);
        //    majorFont1.Append(supplementalFont23);
        //    majorFont1.Append(supplementalFont24);
        //    majorFont1.Append(supplementalFont25);
        //    majorFont1.Append(supplementalFont26);
        //    majorFont1.Append(supplementalFont27);
        //    majorFont1.Append(supplementalFont28);
        //    majorFont1.Append(supplementalFont29);
        //    majorFont1.Append(supplementalFont30);

        //    A.MinorFont minorFont1 = new A.MinorFont();
        //    A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
        //    A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
        //    A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
        //    A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
        //    A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
        //    A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
        //    A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
        //    A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
        //    A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
        //    A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
        //    A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
        //    A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
        //    A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
        //    A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
        //    A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
        //    A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
        //    A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
        //    A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
        //    A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
        //    A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
        //    A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
        //    A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
        //    A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
        //    A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
        //    A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
        //    A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
        //    A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
        //    A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
        //    A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
        //    A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
        //    A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
        //    A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
        //    A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

        //    minorFont1.Append(latinFont2);
        //    minorFont1.Append(eastAsianFont2);
        //    minorFont1.Append(complexScriptFont2);
        //    minorFont1.Append(supplementalFont31);
        //    minorFont1.Append(supplementalFont32);
        //    minorFont1.Append(supplementalFont33);
        //    minorFont1.Append(supplementalFont34);
        //    minorFont1.Append(supplementalFont35);
        //    minorFont1.Append(supplementalFont36);
        //    minorFont1.Append(supplementalFont37);
        //    minorFont1.Append(supplementalFont38);
        //    minorFont1.Append(supplementalFont39);
        //    minorFont1.Append(supplementalFont40);
        //    minorFont1.Append(supplementalFont41);
        //    minorFont1.Append(supplementalFont42);
        //    minorFont1.Append(supplementalFont43);
        //    minorFont1.Append(supplementalFont44);
        //    minorFont1.Append(supplementalFont45);
        //    minorFont1.Append(supplementalFont46);
        //    minorFont1.Append(supplementalFont47);
        //    minorFont1.Append(supplementalFont48);
        //    minorFont1.Append(supplementalFont49);
        //    minorFont1.Append(supplementalFont50);
        //    minorFont1.Append(supplementalFont51);
        //    minorFont1.Append(supplementalFont52);
        //    minorFont1.Append(supplementalFont53);
        //    minorFont1.Append(supplementalFont54);
        //    minorFont1.Append(supplementalFont55);
        //    minorFont1.Append(supplementalFont56);
        //    minorFont1.Append(supplementalFont57);
        //    minorFont1.Append(supplementalFont58);
        //    minorFont1.Append(supplementalFont59);
        //    minorFont1.Append(supplementalFont60);

        //    fontScheme2.Append(majorFont1);
        //    fontScheme2.Append(minorFont1);

        //    A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Стандартная" };

        //    A.FillStyleList fillStyleList1 = new A.FillStyleList();

        //    A.SolidFill solidFill1 = new A.SolidFill();
        //    A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        //    solidFill1.Append(schemeColor1);

        //    A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

        //    A.GradientStopList gradientStopList1 = new A.GradientStopList();

        //    A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

        //    A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint1 = new A.Tint() { Val = 50000 };
        //    A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

        //    schemeColor2.Append(tint1);
        //    schemeColor2.Append(saturationModulation1);

        //    gradientStop1.Append(schemeColor2);

        //    A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

        //    A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint2 = new A.Tint() { Val = 37000 };
        //    A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

        //    schemeColor3.Append(tint2);
        //    schemeColor3.Append(saturationModulation2);

        //    gradientStop2.Append(schemeColor3);

        //    A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

        //    A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint3 = new A.Tint() { Val = 15000 };
        //    A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

        //    schemeColor4.Append(tint3);
        //    schemeColor4.Append(saturationModulation3);

        //    gradientStop3.Append(schemeColor4);

        //    gradientStopList1.Append(gradientStop1);
        //    gradientStopList1.Append(gradientStop2);
        //    gradientStopList1.Append(gradientStop3);
        //    A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

        //    gradientFill1.Append(gradientStopList1);
        //    gradientFill1.Append(linearGradientFill1);

        //    A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

        //    A.GradientStopList gradientStopList2 = new A.GradientStopList();

        //    A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

        //    A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade1 = new A.Shade() { Val = 51000 };
        //    A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

        //    schemeColor5.Append(shade1);
        //    schemeColor5.Append(saturationModulation4);

        //    gradientStop4.Append(schemeColor5);

        //    A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

        //    A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade2 = new A.Shade() { Val = 93000 };
        //    A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

        //    schemeColor6.Append(shade2);
        //    schemeColor6.Append(saturationModulation5);

        //    gradientStop5.Append(schemeColor6);

        //    A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

        //    A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade3 = new A.Shade() { Val = 94000 };
        //    A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

        //    schemeColor7.Append(shade3);
        //    schemeColor7.Append(saturationModulation6);

        //    gradientStop6.Append(schemeColor7);

        //    gradientStopList2.Append(gradientStop4);
        //    gradientStopList2.Append(gradientStop5);
        //    gradientStopList2.Append(gradientStop6);
        //    A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

        //    gradientFill2.Append(gradientStopList2);
        //    gradientFill2.Append(linearGradientFill2);

        //    fillStyleList1.Append(solidFill1);
        //    fillStyleList1.Append(gradientFill1);
        //    fillStyleList1.Append(gradientFill2);

        //    A.LineStyleList lineStyleList1 = new A.LineStyleList();

        //    A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        //    A.SolidFill solidFill2 = new A.SolidFill();

        //    A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade4 = new A.Shade() { Val = 95000 };
        //    A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

        //    schemeColor8.Append(shade4);
        //    schemeColor8.Append(saturationModulation7);

        //    solidFill2.Append(schemeColor8);
        //    A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

        //    outline1.Append(solidFill2);
        //    outline1.Append(presetDash1);

        //    A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        //    A.SolidFill solidFill3 = new A.SolidFill();
        //    A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        //    solidFill3.Append(schemeColor9);
        //    A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

        //    outline2.Append(solidFill3);
        //    outline2.Append(presetDash2);

        //    A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

        //    A.SolidFill solidFill4 = new A.SolidFill();
        //    A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        //    solidFill4.Append(schemeColor10);
        //    A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

        //    outline3.Append(solidFill4);
        //    outline3.Append(presetDash3);

        //    lineStyleList1.Append(outline1);
        //    lineStyleList1.Append(outline2);
        //    lineStyleList1.Append(outline3);

        //    A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

        //    A.EffectStyle effectStyle1 = new A.EffectStyle();

        //    A.EffectList effectList1 = new A.EffectList();

        //    A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

        //    A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
        //    A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

        //    rgbColorModelHex11.Append(alpha1);

        //    outerShadow1.Append(rgbColorModelHex11);

        //    effectList1.Append(outerShadow1);

        //    effectStyle1.Append(effectList1);

        //    A.EffectStyle effectStyle2 = new A.EffectStyle();

        //    A.EffectList effectList2 = new A.EffectList();

        //    A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

        //    A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
        //    A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

        //    rgbColorModelHex12.Append(alpha2);

        //    outerShadow2.Append(rgbColorModelHex12);

        //    effectList2.Append(outerShadow2);

        //    effectStyle2.Append(effectList2);

        //    A.EffectStyle effectStyle3 = new A.EffectStyle();

        //    A.EffectList effectList3 = new A.EffectList();

        //    A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

        //    A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
        //    A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

        //    rgbColorModelHex13.Append(alpha3);

        //    outerShadow3.Append(rgbColorModelHex13);

        //    effectList3.Append(outerShadow3);

        //    A.Scene3DType scene3DType1 = new A.Scene3DType();

        //    A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
        //    A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

        //    camera1.Append(rotation1);

        //    A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
        //    A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

        //    lightRig1.Append(rotation2);

        //    scene3DType1.Append(camera1);
        //    scene3DType1.Append(lightRig1);

        //    A.Shape3DType shape3DType1 = new A.Shape3DType();
        //    A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

        //    shape3DType1.Append(bevelTop1);

        //    effectStyle3.Append(effectList3);
        //    effectStyle3.Append(scene3DType1);
        //    effectStyle3.Append(shape3DType1);

        //    effectStyleList1.Append(effectStyle1);
        //    effectStyleList1.Append(effectStyle2);
        //    effectStyleList1.Append(effectStyle3);

        //    A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

        //    A.SolidFill solidFill5 = new A.SolidFill();
        //    A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

        //    solidFill5.Append(schemeColor11);

        //    A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

        //    A.GradientStopList gradientStopList3 = new A.GradientStopList();

        //    A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

        //    A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint4 = new A.Tint() { Val = 40000 };
        //    A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

        //    schemeColor12.Append(tint4);
        //    schemeColor12.Append(saturationModulation8);

        //    gradientStop7.Append(schemeColor12);

        //    A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

        //    A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint5 = new A.Tint() { Val = 45000 };
        //    A.Shade shade5 = new A.Shade() { Val = 99000 };
        //    A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

        //    schemeColor13.Append(tint5);
        //    schemeColor13.Append(shade5);
        //    schemeColor13.Append(saturationModulation9);

        //    gradientStop8.Append(schemeColor13);

        //    A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

        //    A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade6 = new A.Shade() { Val = 20000 };
        //    A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

        //    schemeColor14.Append(shade6);
        //    schemeColor14.Append(saturationModulation10);

        //    gradientStop9.Append(schemeColor14);

        //    gradientStopList3.Append(gradientStop7);
        //    gradientStopList3.Append(gradientStop8);
        //    gradientStopList3.Append(gradientStop9);

        //    A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
        //    A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

        //    pathGradientFill1.Append(fillToRectangle1);

        //    gradientFill3.Append(gradientStopList3);
        //    gradientFill3.Append(pathGradientFill1);

        //    A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

        //    A.GradientStopList gradientStopList4 = new A.GradientStopList();

        //    A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

        //    A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Tint tint6 = new A.Tint() { Val = 80000 };
        //    A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

        //    schemeColor15.Append(tint6);
        //    schemeColor15.Append(saturationModulation11);

        //    gradientStop10.Append(schemeColor15);

        //    A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

        //    A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
        //    A.Shade shade7 = new A.Shade() { Val = 30000 };
        //    A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

        //    schemeColor16.Append(shade7);
        //    schemeColor16.Append(saturationModulation12);

        //    gradientStop11.Append(schemeColor16);

        //    gradientStopList4.Append(gradientStop10);
        //    gradientStopList4.Append(gradientStop11);

        //    A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
        //    A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

        //    pathGradientFill2.Append(fillToRectangle2);

        //    gradientFill4.Append(gradientStopList4);
        //    gradientFill4.Append(pathGradientFill2);

        //    backgroundFillStyleList1.Append(solidFill5);
        //    backgroundFillStyleList1.Append(gradientFill3);
        //    backgroundFillStyleList1.Append(gradientFill4);

        //    formatScheme1.Append(fillStyleList1);
        //    formatScheme1.Append(lineStyleList1);
        //    formatScheme1.Append(effectStyleList1);
        //    formatScheme1.Append(backgroundFillStyleList1);

        //    themeElements1.Append(colorScheme1);
        //    themeElements1.Append(fontScheme2);
        //    themeElements1.Append(formatScheme1);
        //    A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
        //    A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

        //    theme1.Append(themeElements1);
        //    theme1.Append(objectDefaults1);
        //    theme1.Append(extraColorSchemeList1);

        //    themePart1.Theme = theme1;
        //}

        // Generates content of worksheetPart.
        private static void GenerateWorksheetPartContent(WorksheetPart worksheetPart)
        {
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:C4" };
            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
            SheetData sheetData = new SheetData();

            foreach (CellData d in cellsData)
            {
                int i = (int)d.Row;
                int j = (int)d.Column;
                int k = (int)d.Styleindex;

                Row row = GetRow(sheetData, i);
                Cell cell = new Cell() { CellReference = GetExcelColumnName(j) + i, StyleIndex = (UInt32Value)(UInt32)k };

                SetFormatedCellData(cell, d.Data.ToString(), i, j);
                InsertCellIntoRow(cell, row);
            }

            Columns columns = new Columns();
            InsertColumnWidth(columns);
            MergeCells mergeCells = new MergeCells() { Count = (UInt32Value)1U };
            SetMergeCell(mergeCells);

            SheetViews sheetViews = new SheetViews();
            sheetViews.Append(new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U });
            worksheet.Append(new SheetDimension() { Reference = "A1" });
            worksheet.Append(sheetViews);

            worksheet.Append(new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D });
            if (columnWidthArr.Count > 0)
                worksheet.Append(columns);
            worksheet.Append(sheetData);
            if (mergeArr.Count > 0)
                worksheet.Append(mergeCells);
            worksheetPart.Worksheet = worksheet;
            worksheet.Save();

        }

        private static void SetFormatedCellData(Cell cell, string data, int i, int j)
        {

            if (data == null || data.Length == 0)
                return;

            int number;
            decimal dec;
            CellValue cellValue = new CellValue();
            CellFormula cellFormula = new CellFormula();
            if (data.Substring(0, 1) == "=")
            {

                cell.DataType = new EnumValue<CellValues>(CellValues.String);

                if (data.ToString().Contains('R') && data.ToString().Contains('C'))
                {
                    string convertedFormula = TransformFormulaToA1Format(data, i, GetExcelColumnName(j));
                    cellFormula.Text = convertedFormula.Remove(0, 1);
                    cell.Append(cellFormula);
                }
                else
                {
                    cellFormula.Text = data.Remove(0, 1);
                    cell.Append(cellFormula);
                }


            }
            else if (Int32.TryParse(data, out number))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cellValue.Text = number.ToString();
                cell.Append(cellValue);
            }
            else if (Decimal.TryParse(data, NumberStyles.Number, CultureInfo.InstalledUICulture, out dec))
            {
                string NumberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;
                string correctString = data.Replace(NumberDecimalSeparator, ".");
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cellValue.Text = correctString;
                cell.Append(cellValue);
            }
            else
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cellValue.Text = data.ToString();
                cell.Append(cellValue);
            }
        }

        // Generates content of sharedStringTablePart1.
        private static void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable();
            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private static void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Description = "Created using OpenXML SDK and .NET 3.5";
            document.PackageProperties.Version = "1.0.0.1";
            document.PackageProperties.Creator = "ORC";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime(DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ"), System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddColumnWidth(int columnIndex, double columnWidth)
        {
            columnWidthArr[columnIndex] = columnWidth;
            return 0;
        }
        private static void InsertColumnWidth(Columns columns)
        {

            foreach (var d in columnWidthArr)
            {

                Column column = new Column()
                {
                    Min = (UInt32Value)(UInt32)d.Key,
                    Max = (UInt32Value)(UInt32)d.Key,
                    Width = d.Value,
                    CustomWidth = true
                };
                columns.Append(column);
            }
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddRowHeight(int rowIndex, double rowHeight)
        {
            rowHeightArr[rowIndex] = rowHeight;
            return 0;
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddMergeCell(string mergedCellsId)
        {
            mergeArr.Add(mergedCellsId);
            return 0;
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddCellData(
            int rowIndex,
            int colIndex,
            string data,
            bool bold,
            int size,
            bool wrapText,
            int lineStyle,
            string fontName,
            bool italic,
            bool underline,
            int horizontalAlignment,
            int verticalAlignment,
            bool treead,
            bool grid)
        {
            int ffi = SetFontFormat(bold, size, fontName, italic, underline);
            int cfi = SetCellFormat(wrapText, lineStyle, horizontalAlignment, verticalAlignment, treead, grid, ffi);
            cellsData.Add(new CellData(rowIndex, colIndex, data, cfi));





            return 0;
        }

        private static Row GetRow(SheetData sheetData, int rowIndex)
        {
            Row newRow = null;
            foreach (Row current in sheetData.Elements<Row>())
            {
                if (current.RowIndex >= rowIndex)
                {
                    if (current.RowIndex == rowIndex)
                        return current;
                    newRow = GetNewRow(rowIndex);
                    sheetData.InsertBefore<Row>(newRow, current);
                    return newRow;
                }
            }
            newRow = GetNewRow(rowIndex);
            sheetData.Append(newRow);
            return newRow;
        }

        private static Row GetNewRow(int i)
        {
            double? rowHeight;
            Row row = null;
            if (rowHeightArr.TryGetValue(i, out rowHeight))
            {
                row = new Row()
                {
                    RowIndex = ((UInt32Value)(UInt32)i),
                    Height = rowHeight,
                    CustomHeight = true,
                    Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                    DyDescent = 0.25D
                };
            }
            else
            {
                row = new Row()
                {
                    RowIndex = ((UInt32Value)(UInt32)i),
                    Spans = new ListValue<StringValue>() { InnerText = "1:3" },
                    DyDescent = 0.25D
                };
            }
            return row;
        }
        private static void InsertCellIntoRow(Cell cell, Row row)
        {
            int cellColIndex = GetExcelColumnNumber(cell.CellReference.ToString());
            foreach (Cell current in row.Elements<Cell>())
            {
                int comp = GetExcelColumnNumber(current.CellReference.ToString()) - cellColIndex;
                if (comp >= 0)
                {
                    if (comp == 0)
                    {
                        row.InsertBefore<Cell>(cell, current);
                        current.Remove();
                        return;
                    }
                    row.InsertBefore<Cell>(cell, current);
                    return;
                }
            }
            row.Append(cell);
            return;
        }

        private static void SetMergeCell(MergeCells mergeCells)
        {
            foreach (string d in mergeArr)
            {
                MergeCell mergeCell = new MergeCell() { Reference = d };
                mergeCells.Append(mergeCell);
            }

            PageMargins pageMargins = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
        }


        private static int SetCellFormat(
            bool wrapText,
            int lineStyle,
            int horizontalAlignment,
            int verticalAlignment,
            bool treead,
            bool grid,
            int fontIndex)
        {
            CellStyleFormat currentCell = new CellStyleFormat(wrapText, lineStyle, horizontalAlignment, verticalAlignment, treead, grid, fontIndex);

            int counter = 1;
            foreach (CellStyleFormat d in cellStyleFormatList)
            {
                if (currentCell.Equals(d))
                    return counter;
                counter++;
            }
            cellStyleFormatList.Add(currentCell);
            return cellStyleFormatList.Count;

        }
        private static int SetFontFormat(bool bold, int size, string fontName, bool italic, bool underline)
        {
            FontStyleFormat currentFont = new FontStyleFormat(bold, size, fontName, italic, underline);

            int fontCounter = 1;
            foreach (FontStyleFormat f in fontStyleFormatList)
            {
                if (currentFont.Equals(f))
                    return fontCounter;
                fontCounter++;
            }
            fontStyleFormatList.Add(currentFont);
            return fontStyleFormatList.Count;
        }

        private static long ClearArray()
        {
            cellsData.Clear();
            rowHeightArr.Clear();
            columnWidthArr.Clear();
            mergeArr.Clear();
            fontStyleFormatList.Clear();
            cellStyleFormatList.Clear();
            return 0;
        }


        private static string TransformFormulaToA1Format(string formula, int row, string column)
        {

            String pattern = @"((R((\d*)|(\[[-+]?\d*\])))?C((\d+)|(\[[-+]?\d+\])))"
                                + @"|(R((\d+)|(\[[-+]?\d+\]))(C((\d*)|(\[[-+]?\d*\])))?)";
            return Regex.Replace(formula, pattern, new MatchEvaluator(delegate(Match match) { return TransformReferenceToA1Format(match, row, column); }));
        }



        private static string TransformReferenceToA1Format(Match match, int row, string column)
        {
            string reference = match.Value;

            string rowRefPattern = @"R\[([-+]?\d*)\]";
            string rowAbsPattern = @"R(\d+)";
            string colRefPattern = @"C\[([-+]?\d*)\]";
            string colAbsPattern = @"C(\d+)";


            Boolean isRowRef = Regex.IsMatch(reference, rowRefPattern);
            Boolean isRowAbs = Regex.IsMatch(reference, rowAbsPattern);
            Boolean isColRef = Regex.IsMatch(reference, colRefPattern);
            Boolean isColAbs = Regex.IsMatch(reference, colAbsPattern);
            Boolean hasRow = Regex.IsMatch(reference, "R");
            Boolean hasCol = Regex.IsMatch(reference, "C");


            string currentRowPattern = null;
            string currentColPattern = null;
            string result = "";

            if (hasCol && !(isColAbs || isColRef))
                result = column;

            else if (isColAbs)
                currentColPattern = colAbsPattern;
            else
                currentColPattern = colRefPattern;


            if (currentColPattern != null && hasCol)
            {

                MatchCollection mc = Regex.Matches(reference, currentColPattern);
                if (isColAbs)
                    column = GetExcelColumnName(Int32.Parse(mc[0].Groups[1].Value));

                else
                    column = GetExcelColumnName(Int32.Parse(mc[0].Groups[1].Value) + GetExcelColumnNumber(column));


                if (hasCol && !hasRow)
                    return column + ":" + column;
                else
                    result = column;
            }


            if (hasRow && !(isRowAbs || isRowRef))
                result += row;

            else if (isRowAbs)
                currentRowPattern = rowAbsPattern;
            else
                currentRowPattern = rowRefPattern;


            if (currentRowPattern != null)
            {
                MatchCollection mc = Regex.Matches(reference, currentRowPattern);
                if (isRowAbs)
                    row = Int32.Parse(mc[0].Groups[1].Value);
                else

                    row = (Int32.Parse(mc[0].Groups[1].Value) + row);


                if (hasRow && !hasCol)
                    return row + ":" + row;
                else
                    result += row;
            }


            return result;
        }

    }
}
public class CellData
{
    int rowIndex;
    int colIndex;
    string data;
    int styleIndex;
    public CellData(int rowIndex, int colIndex, string data, int styleIndex)
    {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.data = data;
        this.styleIndex = styleIndex;
    }
    public int Row
    {
        get { return rowIndex; }
        set { rowIndex = value; }
    }
    public int Column
    {
        get { return colIndex; }
        set { colIndex = value; }
    }
    public string Data
    {
        get { return data; }
        set { data = value; }
    }
    public int Styleindex
    {
        get { return styleIndex; }
        set { styleIndex = value; }
    }
}


class CellStyleFormat
{
    bool wrapText;
    int lineStyle;
    int horizontalAlignment;
    int verticalAlignment;
    bool treead;
    bool grid;
    int fontIndex;
    public CellStyleFormat(
        bool wrapText,
        int lineStyle,
        int horizontalAlignment,
        int verticalAlignment,
        bool treead,
        bool grid,
        int fontIndex)
    {
        this.wrapText = wrapText;
        this.lineStyle = lineStyle;
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
        this.treead = treead;
        this.grid = grid;
        this.fontIndex = fontIndex;
    }


    public bool WrapText
    {
        get { return wrapText; }
        set { wrapText = value; }
    }
    public int LineStyle
    {
        get { return lineStyle; }
        set { lineStyle = value; }
    }

    public int HorizontalAlignment
    {
        get { return horizontalAlignment; }
        set { horizontalAlignment = value; }
    }
    public int VerticalAlignment
    {
        get { return verticalAlignment; }
        set { verticalAlignment = value; }
    }
    public bool Treead
    {
        get { return treead; }
        set { treead = value; }
    }
    public bool Grid
    {
        get { return grid; }
        set { grid = value; }
    }
    public int FontIndex
    {
        get { return fontIndex; }
        set { fontIndex = value; }
    }
    public override bool Equals(object obj)
    {
        CellStyleFormat cfs = obj as CellStyleFormat;
        if (cfs == null)
            return false;
        return this.WrapText == cfs.wrapText && this.LineStyle == cfs.lineStyle && this.HorizontalAlignment == cfs.horizontalAlignment && this.VerticalAlignment == cfs.verticalAlignment && this.Treead == cfs.treead && this.Grid == cfs.grid && this.fontIndex == cfs.fontIndex;
    }
}


class FontStyleFormat
{
    bool bold;
    double size;
    string fontName;
    bool italic;
    bool underline;
    public FontStyleFormat(
        bool bold,
        double size,
        string fontName,
        bool italic,
        bool underline)
    {
        this.bold = bold;
        this.size = size;
        this.fontName = fontName;
        this.italic = italic;
        this.underline = underline;

    }

    public bool Bold
    {
        get { return bold; }
        set { bold = value; }
    }
    public double Size
    {
        get { return size; }
        set { size = value; }
    }

    public string FontName
    {
        get { return fontName; }
        set { fontName = value; }
    }
    public bool Italic
    {
        get { return italic; }
        set { italic = value; }
    }
    public bool Underline
    {
        get { return underline; }
        set { underline = value; }
    }
    public override bool Equals(object obj)
    {
        FontStyleFormat ffs = obj as FontStyleFormat;
        if (ffs == null)
            return false;
        return this.Bold == ffs.bold && this.Size == ffs.size && this.FontName.CompareTo(ffs.fontName) == 0 && this.Italic == ffs.italic && this.Underline == ffs.underline;
    }


}


// boolean ib_setvisible = true //При разрушении объекта сделать визуальным
//boolean ib_setprint = false //При разрушении объекта отправить на печать
//integer   ii_Orientation = 1 //Ориентация (1- книжная, 2- альбомная)
//boolean ib_Zoom = false //
//long il_startrow = 1 // Строка с которой начинается вывод таблицы
//integer ii_setFitToPagesWide = 1 //Разместить не более чем на 1 странице в ширину
//integer ii_setFitToPagesTall =  1 //Разместить не более чем на 1 страниц в высоту
//integer ii_nozms = 0 //Обрезать не значащие нули вместе с разделителем


////Параметры шрифта и ячейки, применяются для каждой выводимой ячейки
//boolean ib_setBold = false //Жирный
//integer ii_setsize = 12 //Размер шрифта
//string is_setFontName = "Times New Roman" 
//boolean ib_setItalic = False //Курсив
//boolean ib_setUnderline = False //Подчеркнутый

//boolean ib_setWrapText = False  //Переносить по словам
//integer ii_setLineStyle = 0 //Границы
// integer ii_setHorizontalAlignment = 1 //Выравнивание по горизонтали (3 - по центру)
// integer ii_setVerticalAlignment = 2 //Выравнивание по вертикали (2- по центру)
// boolean ib_treead = false //Разделять на триады числовые значения
// boolean ib_topline  = false //Верхняя граница
// boolean ib_grid  = false //Видимость сетки


////---------
//string is_fn = 'c:\temp\temp.xls'
//boolean ib_openfile = true

//// Стиль ссылки
//// Application.ReferenceStyle = xlA1
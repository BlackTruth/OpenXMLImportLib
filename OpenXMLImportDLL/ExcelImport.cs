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

namespace OpenXMLImportDLL
{
    public class ExcelImport
    {
        static SortedList<int, double?> rowHeightArr = new SortedList<int, double?>();
        static ArrayList columnWidthArr = new ArrayList();
        static List<CellData> cellsData = new List<CellData>();

        //Convert Excel column number to Excel column name (1=A, 2=B)
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
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

                ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
                GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

                WorkbookPart workbookPart = document.AddWorkbookPart();
                GenerateWorkbookPartContent(workbookPart);

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
                GenerateWorkbookStylesPartContent(workbookStylesPart); //Стили ячеек

                ThemePart themePart1 = workbookPart.AddNewPart<ThemePart>("rId2");
                GenerateThemePart1Content(themePart1);

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
                GenerateWorksheetPartContent(worksheetPart);

                SharedStringTablePart sharedStringTablePart1 = workbookPart.AddNewPart<SharedStringTablePart>("rId4");
                GenerateSharedStringTablePart1Content(sharedStringTablePart1);

                SetPackageProperties(document);

            }
            return 0;


        }

        //Generates content of extendedFilePropertiesPart1.
        private static void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Листы";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Лист1";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "15.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

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

            CalculationProperties calculationProperties = new CalculationProperties() { CalculationId = (UInt32Value)152511U };
            calculationProperties.ForceFullCalculation = true;
            calculationProperties.FullCalculationOnLoad = true;


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

            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)18U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 18D };
            Color color3 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName3 = new FontName() { Val = "Cambria" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Major };

            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet3);
            font3.Append(fontScheme3);

            Font font4 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 15D };
            Color color4 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(bold1);
            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontCharSet4);
            font4.Append(fontScheme4);

            Font font5 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = 13D };
            Color color5 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName5 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font5.Append(bold2);
            font5.Append(fontSize5);
            font5.Append(color5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);
            font5.Append(fontCharSet5);
            font5.Append(fontScheme5);

            Font font6 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize6 = new FontSize() { Val = 11D };
            Color color6 = new Color() { Theme = (UInt32Value)3U };
            FontName fontName6 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme6 = new FontScheme() { Val = FontSchemeValues.Minor };

            font6.Append(bold3);
            font6.Append(fontSize6);
            font6.Append(color6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);
            font6.Append(fontCharSet6);
            font6.Append(fontScheme6);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 11D };
            Color color7 = new Color() { Rgb = "FF006100" };
            FontName fontName7 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet7 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme7 = new FontScheme() { Val = FontSchemeValues.Minor };

            font7.Append(fontSize7);
            font7.Append(color7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);
            font7.Append(fontCharSet7);
            font7.Append(fontScheme7);

            Font font8 = new Font();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            Color color8 = new Color() { Rgb = "FF9C0006" };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet8 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme8 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(fontSize8);
            font8.Append(color8);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);
            font8.Append(fontCharSet8);
            font8.Append(fontScheme8);

            Font font9 = new Font();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color9 = new Color() { Rgb = "FF9C6500" };
            FontName fontName9 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet9 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme9 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(fontSize9);
            font9.Append(color9);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontCharSet9);
            font9.Append(fontScheme9);

            Font font10 = new Font();
            FontSize fontSize10 = new FontSize() { Val = 11D };
            Color color10 = new Color() { Rgb = "FF3F3F76" };
            FontName fontName10 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet10 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme10 = new FontScheme() { Val = FontSchemeValues.Minor };

            font10.Append(fontSize10);
            font10.Append(color10);
            font10.Append(fontName10);
            font10.Append(fontFamilyNumbering10);
            font10.Append(fontCharSet10);
            font10.Append(fontScheme10);

            Font font11 = new Font();
            Bold bold4 = new Bold();
            FontSize fontSize11 = new FontSize() { Val = 11D };
            Color color11 = new Color() { Rgb = "FF3F3F3F" };
            FontName fontName11 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet11 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme11 = new FontScheme() { Val = FontSchemeValues.Minor };

            font11.Append(bold4);
            font11.Append(fontSize11);
            font11.Append(color11);
            font11.Append(fontName11);
            font11.Append(fontFamilyNumbering11);
            font11.Append(fontCharSet11);
            font11.Append(fontScheme11);

            Font font12 = new Font();
            Bold bold5 = new Bold();
            FontSize fontSize12 = new FontSize() { Val = 11D };
            Color color12 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName12 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet12 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme12 = new FontScheme() { Val = FontSchemeValues.Minor };

            font12.Append(bold5);
            font12.Append(fontSize12);
            font12.Append(color12);
            font12.Append(fontName12);
            font12.Append(fontFamilyNumbering12);
            font12.Append(fontCharSet12);
            font12.Append(fontScheme12);

            Font font13 = new Font();
            FontSize fontSize13 = new FontSize() { Val = 11D };
            Color color13 = new Color() { Rgb = "FFFA7D00" };
            FontName fontName13 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet13 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme13 = new FontScheme() { Val = FontSchemeValues.Minor };

            font13.Append(fontSize13);
            font13.Append(color13);
            font13.Append(fontName13);
            font13.Append(fontFamilyNumbering13);
            font13.Append(fontCharSet13);
            font13.Append(fontScheme13);

            Font font14 = new Font();
            Bold bold6 = new Bold();
            FontSize fontSize14 = new FontSize() { Val = 11D };
            Color color14 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName14 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet14 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme14 = new FontScheme() { Val = FontSchemeValues.Minor };

            font14.Append(bold6);
            font14.Append(fontSize14);
            font14.Append(color14);
            font14.Append(fontName14);
            font14.Append(fontFamilyNumbering14);
            font14.Append(fontCharSet14);
            font14.Append(fontScheme14);

            Font font15 = new Font();
            FontSize fontSize15 = new FontSize() { Val = 11D };
            Color color15 = new Color() { Rgb = "FFFF0000" };
            FontName fontName15 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering15 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet15 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme15 = new FontScheme() { Val = FontSchemeValues.Minor };

            font15.Append(fontSize15);
            font15.Append(color15);
            font15.Append(fontName15);
            font15.Append(fontFamilyNumbering15);
            font15.Append(fontCharSet15);
            font15.Append(fontScheme15);

            Font font16 = new Font();
            Italic italic1 = new Italic();
            FontSize fontSize16 = new FontSize() { Val = 11D };
            Color color16 = new Color() { Rgb = "FF7F7F7F" };
            FontName fontName16 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering16 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet16 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme16 = new FontScheme() { Val = FontSchemeValues.Minor };

            font16.Append(italic1);
            font16.Append(fontSize16);
            font16.Append(color16);
            font16.Append(fontName16);
            font16.Append(fontFamilyNumbering16);
            font16.Append(fontCharSet16);
            font16.Append(fontScheme16);

            Font font17 = new Font();
            Bold bold7 = new Bold();
            FontSize fontSize17 = new FontSize() { Val = 11D };
            Color color17 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName17 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering17 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet17 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme17 = new FontScheme() { Val = FontSchemeValues.Minor };

            font17.Append(bold7);
            font17.Append(fontSize17);
            font17.Append(color17);
            font17.Append(fontName17);
            font17.Append(fontFamilyNumbering17);
            font17.Append(fontCharSet17);
            font17.Append(fontScheme17);

            Font font18 = new Font();
            FontSize fontSize18 = new FontSize() { Val = 11D };
            Color color18 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName18 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering18 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet18 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme18 = new FontScheme() { Val = FontSchemeValues.Minor };

            font18.Append(fontSize18);
            font18.Append(color18);
            font18.Append(fontName18);
            font18.Append(fontFamilyNumbering18);
            font18.Append(fontCharSet18);
            font18.Append(fontScheme18);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);
            fonts1.Append(font11);
            fonts1.Append(font12);
            fonts1.Append(font13);
            fonts1.Append(font14);
            fonts1.Append(font15);
            fonts1.Append(font16);
            fonts1.Append(font17);
            fonts1.Append(font18);

            Fills fills1 = new Fills() { Count = (UInt32Value)33U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFC6EFCE" };

            patternFill3.Append(foregroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFFC7CE" };

            patternFill4.Append(foregroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFEB9C" };

            patternFill5.Append(foregroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FFFFCC99" };

            patternFill6.Append(foregroundColor4);

            fill6.Append(patternFill6);

            Fill fill7 = new Fill();

            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "FFF2F2F2" };

            patternFill7.Append(foregroundColor5);

            fill7.Append(patternFill7);

            Fill fill8 = new Fill();

            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFA5A5A5" };

            patternFill8.Append(foregroundColor6);

            fill8.Append(patternFill8);

            Fill fill9 = new Fill();

            PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor7 = new ForegroundColor() { Rgb = "FFFFFFCC" };

            patternFill9.Append(foregroundColor7);

            fill9.Append(patternFill9);

            Fill fill10 = new Fill();

            PatternFill patternFill10 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor8 = new ForegroundColor() { Theme = (UInt32Value)4U };

            patternFill10.Append(foregroundColor8);

            fill10.Append(patternFill10);

            Fill fill11 = new Fill();

            PatternFill patternFill11 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor9 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill11.Append(foregroundColor9);
            patternFill11.Append(backgroundColor1);

            fill11.Append(patternFill11);

            Fill fill12 = new Fill();

            PatternFill patternFill12 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor10 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill12.Append(foregroundColor10);
            patternFill12.Append(backgroundColor2);

            fill12.Append(patternFill12);

            Fill fill13 = new Fill();

            PatternFill patternFill13 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor11 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill13.Append(foregroundColor11);
            patternFill13.Append(backgroundColor3);

            fill13.Append(patternFill13);

            Fill fill14 = new Fill();

            PatternFill patternFill14 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor12 = new ForegroundColor() { Theme = (UInt32Value)5U };

            patternFill14.Append(foregroundColor12);

            fill14.Append(patternFill14);

            Fill fill15 = new Fill();

            PatternFill patternFill15 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor13 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill15.Append(foregroundColor13);
            patternFill15.Append(backgroundColor4);

            fill15.Append(patternFill15);

            Fill fill16 = new Fill();

            PatternFill patternFill16 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor14 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill16.Append(foregroundColor14);
            patternFill16.Append(backgroundColor5);

            fill16.Append(patternFill16);

            Fill fill17 = new Fill();

            PatternFill patternFill17 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor15 = new ForegroundColor() { Theme = (UInt32Value)5U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill17.Append(foregroundColor15);
            patternFill17.Append(backgroundColor6);

            fill17.Append(patternFill17);

            Fill fill18 = new Fill();

            PatternFill patternFill18 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor16 = new ForegroundColor() { Theme = (UInt32Value)6U };

            patternFill18.Append(foregroundColor16);

            fill18.Append(patternFill18);

            Fill fill19 = new Fill();

            PatternFill patternFill19 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor17 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill19.Append(foregroundColor17);
            patternFill19.Append(backgroundColor7);

            fill19.Append(patternFill19);

            Fill fill20 = new Fill();

            PatternFill patternFill20 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor18 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill20.Append(foregroundColor18);
            patternFill20.Append(backgroundColor8);

            fill20.Append(patternFill20);

            Fill fill21 = new Fill();

            PatternFill patternFill21 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor19 = new ForegroundColor() { Theme = (UInt32Value)6U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill21.Append(foregroundColor19);
            patternFill21.Append(backgroundColor9);

            fill21.Append(patternFill21);

            Fill fill22 = new Fill();

            PatternFill patternFill22 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor20 = new ForegroundColor() { Theme = (UInt32Value)7U };

            patternFill22.Append(foregroundColor20);

            fill22.Append(patternFill22);

            Fill fill23 = new Fill();

            PatternFill patternFill23 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor21 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor10 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill23.Append(foregroundColor21);
            patternFill23.Append(backgroundColor10);

            fill23.Append(patternFill23);

            Fill fill24 = new Fill();

            PatternFill patternFill24 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor22 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor11 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill24.Append(foregroundColor22);
            patternFill24.Append(backgroundColor11);

            fill24.Append(patternFill24);

            Fill fill25 = new Fill();

            PatternFill patternFill25 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor23 = new ForegroundColor() { Theme = (UInt32Value)7U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor12 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill25.Append(foregroundColor23);
            patternFill25.Append(backgroundColor12);

            fill25.Append(patternFill25);

            Fill fill26 = new Fill();

            PatternFill patternFill26 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor24 = new ForegroundColor() { Theme = (UInt32Value)8U };

            patternFill26.Append(foregroundColor24);

            fill26.Append(patternFill26);

            Fill fill27 = new Fill();

            PatternFill patternFill27 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor25 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor13 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill27.Append(foregroundColor25);
            patternFill27.Append(backgroundColor13);

            fill27.Append(patternFill27);

            Fill fill28 = new Fill();

            PatternFill patternFill28 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor26 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor14 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill28.Append(foregroundColor26);
            patternFill28.Append(backgroundColor14);

            fill28.Append(patternFill28);

            Fill fill29 = new Fill();

            PatternFill patternFill29 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor27 = new ForegroundColor() { Theme = (UInt32Value)8U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor15 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill29.Append(foregroundColor27);
            patternFill29.Append(backgroundColor15);

            fill29.Append(patternFill29);

            Fill fill30 = new Fill();

            PatternFill patternFill30 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor28 = new ForegroundColor() { Theme = (UInt32Value)9U };

            patternFill30.Append(foregroundColor28);

            fill30.Append(patternFill30);

            Fill fill31 = new Fill();

            PatternFill patternFill31 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor29 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor16 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill31.Append(foregroundColor29);
            patternFill31.Append(backgroundColor16);

            fill31.Append(patternFill31);

            Fill fill32 = new Fill();

            PatternFill patternFill32 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor30 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor17 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill32.Append(foregroundColor30);
            patternFill32.Append(backgroundColor17);

            fill32.Append(patternFill32);

            Fill fill33 = new Fill();

            PatternFill patternFill33 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor31 = new ForegroundColor() { Theme = (UInt32Value)9U, Tint = 0.39997558519241921D };
            BackgroundColor backgroundColor18 = new BackgroundColor() { Indexed = (UInt32Value)65U };

            patternFill33.Append(foregroundColor31);
            patternFill33.Append(backgroundColor18);

            fill33.Append(patternFill33);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);
            fills1.Append(fill7);
            fills1.Append(fill8);
            fills1.Append(fill9);
            fills1.Append(fill10);
            fills1.Append(fill11);
            fills1.Append(fill12);
            fills1.Append(fill13);
            fills1.Append(fill14);
            fills1.Append(fill15);
            fills1.Append(fill16);
            fills1.Append(fill17);
            fills1.Append(fill18);
            fills1.Append(fill19);
            fills1.Append(fill20);
            fills1.Append(fill21);
            fills1.Append(fill22);
            fills1.Append(fill23);
            fills1.Append(fill24);
            fills1.Append(fill25);
            fills1.Append(fill26);
            fills1.Append(fill27);
            fills1.Append(fill28);
            fills1.Append(fill29);
            fills1.Append(fill30);
            fills1.Append(fill31);
            fills1.Append(fill32);
            fills1.Append(fill33);

            Borders borders1 = new Borders() { Count = (UInt32Value)15U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();
            LeftBorder leftBorder2 = new LeftBorder();
            RightBorder rightBorder2 = new RightBorder();
            TopBorder topBorder2 = new TopBorder();

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color19 = new Color() { Theme = (UInt32Value)4U };

            bottomBorder2.Append(color19);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            Border border3 = new Border();
            LeftBorder leftBorder3 = new LeftBorder();
            RightBorder rightBorder3 = new RightBorder();
            TopBorder topBorder3 = new TopBorder();

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thick };
            Color color20 = new Color() { Theme = (UInt32Value)4U, Tint = 0.499984740745262D };

            bottomBorder3.Append(color20);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            Border border4 = new Border();
            LeftBorder leftBorder4 = new LeftBorder();
            RightBorder rightBorder4 = new RightBorder();
            TopBorder topBorder4 = new TopBorder();

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color21 = new Color() { Theme = (UInt32Value)4U, Tint = 0.39997558519241921D };

            bottomBorder4.Append(color21);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            Border border5 = new Border();

            LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color22 = new Color() { Rgb = "FF7F7F7F" };

            leftBorder5.Append(color22);

            RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color23 = new Color() { Rgb = "FF7F7F7F" };

            rightBorder5.Append(color23);

            TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color24 = new Color() { Rgb = "FF7F7F7F" };

            topBorder5.Append(color24);

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color25 = new Color() { Rgb = "FF7F7F7F" };

            bottomBorder5.Append(color25);
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            Border border6 = new Border();

            LeftBorder leftBorder6 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color26 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder6.Append(color26);

            RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color27 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder6.Append(color27);

            TopBorder topBorder6 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color28 = new Color() { Rgb = "FF3F3F3F" };

            topBorder6.Append(color28);

            BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color29 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder6.Append(color29);
            DiagonalBorder diagonalBorder6 = new DiagonalBorder();

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);
            border6.Append(diagonalBorder6);

            Border border7 = new Border();
            LeftBorder leftBorder7 = new LeftBorder();
            RightBorder rightBorder7 = new RightBorder();
            TopBorder topBorder7 = new TopBorder();

            BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color30 = new Color() { Rgb = "FFFF8001" };

            bottomBorder7.Append(color30);
            DiagonalBorder diagonalBorder7 = new DiagonalBorder();

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);
            border7.Append(diagonalBorder7);

            Border border8 = new Border();

            LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Double };
            Color color31 = new Color() { Rgb = "FF3F3F3F" };

            leftBorder8.Append(color31);

            RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Double };
            Color color32 = new Color() { Rgb = "FF3F3F3F" };

            rightBorder8.Append(color32);

            TopBorder topBorder8 = new TopBorder() { Style = BorderStyleValues.Double };
            Color color33 = new Color() { Rgb = "FF3F3F3F" };

            topBorder8.Append(color33);

            BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color34 = new Color() { Rgb = "FF3F3F3F" };

            bottomBorder8.Append(color34);
            DiagonalBorder diagonalBorder8 = new DiagonalBorder();

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);
            border8.Append(diagonalBorder8);

            Border border9 = new Border();

            LeftBorder leftBorder9 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color35 = new Color() { Rgb = "FFB2B2B2" };

            leftBorder9.Append(color35);

            RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color36 = new Color() { Rgb = "FFB2B2B2" };

            rightBorder9.Append(color36);

            TopBorder topBorder9 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color37 = new Color() { Rgb = "FFB2B2B2" };

            topBorder9.Append(color37);

            BottomBorder bottomBorder9 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color38 = new Color() { Rgb = "FFB2B2B2" };

            bottomBorder9.Append(color38);
            DiagonalBorder diagonalBorder9 = new DiagonalBorder();

            border9.Append(leftBorder9);
            border9.Append(rightBorder9);
            border9.Append(topBorder9);
            border9.Append(bottomBorder9);
            border9.Append(diagonalBorder9);

            Border border10 = new Border();
            LeftBorder leftBorder10 = new LeftBorder();
            RightBorder rightBorder10 = new RightBorder();

            TopBorder topBorder10 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color39 = new Color() { Theme = (UInt32Value)4U };

            topBorder10.Append(color39);

            BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Double };
            Color color40 = new Color() { Theme = (UInt32Value)4U };

            bottomBorder10.Append(color40);
            DiagonalBorder diagonalBorder10 = new DiagonalBorder();

            border10.Append(leftBorder10);
            border10.Append(rightBorder10);
            border10.Append(topBorder10);
            border10.Append(bottomBorder10);
            border10.Append(diagonalBorder10);

            Border border11 = new Border();
            LeftBorder leftBorder11 = new LeftBorder();
            RightBorder rightBorder11 = new RightBorder();
            TopBorder topBorder11 = new TopBorder();

            BottomBorder bottomBorder11 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color41 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder11.Append(color41);
            DiagonalBorder diagonalBorder11 = new DiagonalBorder();

            border11.Append(leftBorder11);
            border11.Append(rightBorder11);
            border11.Append(topBorder11);
            border11.Append(bottomBorder11);
            border11.Append(diagonalBorder11);

            Border border12 = new Border();
            LeftBorder leftBorder12 = new LeftBorder();
            RightBorder rightBorder12 = new RightBorder();

            TopBorder topBorder12 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color42 = new Color() { Indexed = (UInt32Value)64U };

            topBorder12.Append(color42);
            BottomBorder bottomBorder12 = new BottomBorder();
            DiagonalBorder diagonalBorder12 = new DiagonalBorder();

            border12.Append(leftBorder12);
            border12.Append(rightBorder12);
            border12.Append(topBorder12);
            border12.Append(bottomBorder12);
            border12.Append(diagonalBorder12);

            Border border13 = new Border();

            LeftBorder leftBorder13 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color43 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder13.Append(color43);
            RightBorder rightBorder13 = new RightBorder();
            TopBorder topBorder13 = new TopBorder();
            BottomBorder bottomBorder13 = new BottomBorder();
            DiagonalBorder diagonalBorder13 = new DiagonalBorder();

            border13.Append(leftBorder13);
            border13.Append(rightBorder13);
            border13.Append(topBorder13);
            border13.Append(bottomBorder13);
            border13.Append(diagonalBorder13);

            Border border14 = new Border();
            LeftBorder leftBorder14 = new LeftBorder();

            RightBorder rightBorder14 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color44 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder14.Append(color44);
            TopBorder topBorder14 = new TopBorder();
            BottomBorder bottomBorder14 = new BottomBorder();
            DiagonalBorder diagonalBorder14 = new DiagonalBorder();

            border14.Append(leftBorder14);
            border14.Append(rightBorder14);
            border14.Append(topBorder14);
            border14.Append(bottomBorder14);
            border14.Append(diagonalBorder14);

            Border border15 = new Border();

            LeftBorder leftBorder15 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color45 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder15.Append(color45);

            RightBorder rightBorder15 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color46 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder15.Append(color46);

            TopBorder topBorder15 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color47 = new Color() { Indexed = (UInt32Value)64U };

            topBorder15.Append(color47);

            BottomBorder bottomBorder15 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color48 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder15.Append(color48);
            DiagonalBorder diagonalBorder15 = new DiagonalBorder();

            border15.Append(leftBorder15);
            border15.Append(rightBorder15);
            border15.Append(topBorder15);
            border15.Append(bottomBorder15);
            border15.Append(diagonalBorder15);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);
            borders1.Append(border6);
            borders1.Append(border7);
            borders1.Append(border8);
            borders1.Append(border9);
            borders1.Append(border10);
            borders1.Append(border11);
            borders1.Append(border12);
            borders1.Append(border13);
            borders1.Append(border14);
            borders1.Append(border15);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)42U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)9U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)4U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)10U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)5U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)11U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)4U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)12U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)6U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)7U, ApplyNumberFormat = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)14U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)8U, BorderId = (UInt32Value)8U, ApplyNumberFormat = false, ApplyFont = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)15U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyFill = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)16U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)9U, ApplyNumberFormat = false, ApplyFill = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)9U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)10U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)11U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)12U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)13U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)14U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)15U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)16U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)17U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)18U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)19U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)20U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)21U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)22U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)23U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)24U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)25U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)26U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)27U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)28U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)29U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)30U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)31U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };
            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)17U, FillId = (UInt32Value)32U, BorderId = (UInt32Value)0U, ApplyNumberFormat = false, ApplyBorder = false, ApplyAlignment = false, ApplyProtection = false };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);
            cellStyleFormats1.Append(cellFormat3);
            cellStyleFormats1.Append(cellFormat4);
            cellStyleFormats1.Append(cellFormat5);
            cellStyleFormats1.Append(cellFormat6);
            cellStyleFormats1.Append(cellFormat7);
            cellStyleFormats1.Append(cellFormat8);
            cellStyleFormats1.Append(cellFormat9);
            cellStyleFormats1.Append(cellFormat10);
            cellStyleFormats1.Append(cellFormat11);
            cellStyleFormats1.Append(cellFormat12);
            cellStyleFormats1.Append(cellFormat13);
            cellStyleFormats1.Append(cellFormat14);
            cellStyleFormats1.Append(cellFormat15);
            cellStyleFormats1.Append(cellFormat16);
            cellStyleFormats1.Append(cellFormat17);
            cellStyleFormats1.Append(cellFormat18);
            cellStyleFormats1.Append(cellFormat19);
            cellStyleFormats1.Append(cellFormat20);
            cellStyleFormats1.Append(cellFormat21);
            cellStyleFormats1.Append(cellFormat22);
            cellStyleFormats1.Append(cellFormat23);
            cellStyleFormats1.Append(cellFormat24);
            cellStyleFormats1.Append(cellFormat25);
            cellStyleFormats1.Append(cellFormat26);
            cellStyleFormats1.Append(cellFormat27);
            cellStyleFormats1.Append(cellFormat28);
            cellStyleFormats1.Append(cellFormat29);
            cellStyleFormats1.Append(cellFormat30);
            cellStyleFormats1.Append(cellFormat31);
            cellStyleFormats1.Append(cellFormat32);
            cellStyleFormats1.Append(cellFormat33);
            cellStyleFormats1.Append(cellFormat34);
            cellStyleFormats1.Append(cellFormat35);
            cellStyleFormats1.Append(cellFormat36);
            cellStyleFormats1.Append(cellFormat37);
            cellStyleFormats1.Append(cellFormat38);
            cellStyleFormats1.Append(cellFormat39);
            cellStyleFormats1.Append(cellFormat40);
            cellStyleFormats1.Append(cellFormat41);
            cellStyleFormats1.Append(cellFormat42);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)7U };
            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat45 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat46 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat47 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat48 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)13U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat49 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)14U, FormatId = (UInt32Value)0U, ApplyBorder = true };

            cellFormats1.Append(cellFormat43);
            cellFormats1.Append(cellFormat44);
            cellFormats1.Append(cellFormat45);
            cellFormats1.Append(cellFormat46);
            cellFormats1.Append(cellFormat47);
            cellFormats1.Append(cellFormat48);
            cellFormats1.Append(cellFormat49);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)42U };
            CellStyle cellStyle1 = new CellStyle() { Name = "20% — акцент1", FormatId = (UInt32Value)19U, BuiltinId = (UInt32Value)30U, CustomBuiltin = true };
            CellStyle cellStyle2 = new CellStyle() { Name = "20% — акцент2", FormatId = (UInt32Value)23U, BuiltinId = (UInt32Value)34U, CustomBuiltin = true };
            CellStyle cellStyle3 = new CellStyle() { Name = "20% — акцент3", FormatId = (UInt32Value)27U, BuiltinId = (UInt32Value)38U, CustomBuiltin = true };
            CellStyle cellStyle4 = new CellStyle() { Name = "20% — акцент4", FormatId = (UInt32Value)31U, BuiltinId = (UInt32Value)42U, CustomBuiltin = true };
            CellStyle cellStyle5 = new CellStyle() { Name = "20% — акцент5", FormatId = (UInt32Value)35U, BuiltinId = (UInt32Value)46U, CustomBuiltin = true };
            CellStyle cellStyle6 = new CellStyle() { Name = "20% — акцент6", FormatId = (UInt32Value)39U, BuiltinId = (UInt32Value)50U, CustomBuiltin = true };
            CellStyle cellStyle7 = new CellStyle() { Name = "40% — акцент1", FormatId = (UInt32Value)20U, BuiltinId = (UInt32Value)31U, CustomBuiltin = true };
            CellStyle cellStyle8 = new CellStyle() { Name = "40% — акцент2", FormatId = (UInt32Value)24U, BuiltinId = (UInt32Value)35U, CustomBuiltin = true };
            CellStyle cellStyle9 = new CellStyle() { Name = "40% — акцент3", FormatId = (UInt32Value)28U, BuiltinId = (UInt32Value)39U, CustomBuiltin = true };
            CellStyle cellStyle10 = new CellStyle() { Name = "40% — акцент4", FormatId = (UInt32Value)32U, BuiltinId = (UInt32Value)43U, CustomBuiltin = true };
            CellStyle cellStyle11 = new CellStyle() { Name = "40% — акцент5", FormatId = (UInt32Value)36U, BuiltinId = (UInt32Value)47U, CustomBuiltin = true };
            CellStyle cellStyle12 = new CellStyle() { Name = "40% — акцент6", FormatId = (UInt32Value)40U, BuiltinId = (UInt32Value)51U, CustomBuiltin = true };
            CellStyle cellStyle13 = new CellStyle() { Name = "60% — акцент1", FormatId = (UInt32Value)21U, BuiltinId = (UInt32Value)32U, CustomBuiltin = true };
            CellStyle cellStyle14 = new CellStyle() { Name = "60% — акцент2", FormatId = (UInt32Value)25U, BuiltinId = (UInt32Value)36U, CustomBuiltin = true };
            CellStyle cellStyle15 = new CellStyle() { Name = "60% — акцент3", FormatId = (UInt32Value)29U, BuiltinId = (UInt32Value)40U, CustomBuiltin = true };
            CellStyle cellStyle16 = new CellStyle() { Name = "60% — акцент4", FormatId = (UInt32Value)33U, BuiltinId = (UInt32Value)44U, CustomBuiltin = true };
            CellStyle cellStyle17 = new CellStyle() { Name = "60% — акцент5", FormatId = (UInt32Value)37U, BuiltinId = (UInt32Value)48U, CustomBuiltin = true };
            CellStyle cellStyle18 = new CellStyle() { Name = "60% — акцент6", FormatId = (UInt32Value)41U, BuiltinId = (UInt32Value)52U, CustomBuiltin = true };
            CellStyle cellStyle19 = new CellStyle() { Name = "Акцент1", FormatId = (UInt32Value)18U, BuiltinId = (UInt32Value)29U, CustomBuiltin = true };
            CellStyle cellStyle20 = new CellStyle() { Name = "Акцент2", FormatId = (UInt32Value)22U, BuiltinId = (UInt32Value)33U, CustomBuiltin = true };
            CellStyle cellStyle21 = new CellStyle() { Name = "Акцент3", FormatId = (UInt32Value)26U, BuiltinId = (UInt32Value)37U, CustomBuiltin = true };
            CellStyle cellStyle22 = new CellStyle() { Name = "Акцент4", FormatId = (UInt32Value)30U, BuiltinId = (UInt32Value)41U, CustomBuiltin = true };
            CellStyle cellStyle23 = new CellStyle() { Name = "Акцент5", FormatId = (UInt32Value)34U, BuiltinId = (UInt32Value)45U, CustomBuiltin = true };
            CellStyle cellStyle24 = new CellStyle() { Name = "Акцент6", FormatId = (UInt32Value)38U, BuiltinId = (UInt32Value)49U, CustomBuiltin = true };
            CellStyle cellStyle25 = new CellStyle() { Name = "Ввод ", FormatId = (UInt32Value)9U, BuiltinId = (UInt32Value)20U, CustomBuiltin = true };
            CellStyle cellStyle26 = new CellStyle() { Name = "Вывод", FormatId = (UInt32Value)10U, BuiltinId = (UInt32Value)21U, CustomBuiltin = true };
            CellStyle cellStyle27 = new CellStyle() { Name = "Вычисление", FormatId = (UInt32Value)11U, BuiltinId = (UInt32Value)22U, CustomBuiltin = true };
            CellStyle cellStyle28 = new CellStyle() { Name = "Заголовок 1", FormatId = (UInt32Value)2U, BuiltinId = (UInt32Value)16U, CustomBuiltin = true };
            CellStyle cellStyle29 = new CellStyle() { Name = "Заголовок 2", FormatId = (UInt32Value)3U, BuiltinId = (UInt32Value)17U, CustomBuiltin = true };
            CellStyle cellStyle30 = new CellStyle() { Name = "Заголовок 3", FormatId = (UInt32Value)4U, BuiltinId = (UInt32Value)18U, CustomBuiltin = true };
            CellStyle cellStyle31 = new CellStyle() { Name = "Заголовок 4", FormatId = (UInt32Value)5U, BuiltinId = (UInt32Value)19U, CustomBuiltin = true };
            CellStyle cellStyle32 = new CellStyle() { Name = "Итог", FormatId = (UInt32Value)17U, BuiltinId = (UInt32Value)25U, CustomBuiltin = true };
            CellStyle cellStyle33 = new CellStyle() { Name = "Контрольная ячейка", FormatId = (UInt32Value)13U, BuiltinId = (UInt32Value)23U, CustomBuiltin = true };
            CellStyle cellStyle34 = new CellStyle() { Name = "Название", FormatId = (UInt32Value)1U, BuiltinId = (UInt32Value)15U, CustomBuiltin = true };
            CellStyle cellStyle35 = new CellStyle() { Name = "Нейтральный", FormatId = (UInt32Value)8U, BuiltinId = (UInt32Value)28U, CustomBuiltin = true };
            CellStyle cellStyle36 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
            CellStyle cellStyle37 = new CellStyle() { Name = "Плохой", FormatId = (UInt32Value)7U, BuiltinId = (UInt32Value)27U, CustomBuiltin = true };
            CellStyle cellStyle38 = new CellStyle() { Name = "Пояснение", FormatId = (UInt32Value)16U, BuiltinId = (UInt32Value)53U, CustomBuiltin = true };
            CellStyle cellStyle39 = new CellStyle() { Name = "Примечание", FormatId = (UInt32Value)15U, BuiltinId = (UInt32Value)10U, CustomBuiltin = true };
            CellStyle cellStyle40 = new CellStyle() { Name = "Связанная ячейка", FormatId = (UInt32Value)12U, BuiltinId = (UInt32Value)24U, CustomBuiltin = true };
            CellStyle cellStyle41 = new CellStyle() { Name = "Текст предупреждения", FormatId = (UInt32Value)14U, BuiltinId = (UInt32Value)11U, CustomBuiltin = true };
            CellStyle cellStyle42 = new CellStyle() { Name = "Хороший", FormatId = (UInt32Value)6U, BuiltinId = (UInt32Value)26U, CustomBuiltin = true };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            cellStyles1.Append(cellStyle3);
            cellStyles1.Append(cellStyle4);
            cellStyles1.Append(cellStyle5);
            cellStyles1.Append(cellStyle6);
            cellStyles1.Append(cellStyle7);
            cellStyles1.Append(cellStyle8);
            cellStyles1.Append(cellStyle9);
            cellStyles1.Append(cellStyle10);
            cellStyles1.Append(cellStyle11);
            cellStyles1.Append(cellStyle12);
            cellStyles1.Append(cellStyle13);
            cellStyles1.Append(cellStyle14);
            cellStyles1.Append(cellStyle15);
            cellStyles1.Append(cellStyle16);
            cellStyles1.Append(cellStyle17);
            cellStyles1.Append(cellStyle18);
            cellStyles1.Append(cellStyle19);
            cellStyles1.Append(cellStyle20);
            cellStyles1.Append(cellStyle21);
            cellStyles1.Append(cellStyle22);
            cellStyles1.Append(cellStyle23);
            cellStyles1.Append(cellStyle24);
            cellStyles1.Append(cellStyle25);
            cellStyles1.Append(cellStyle26);
            cellStyles1.Append(cellStyle27);
            cellStyles1.Append(cellStyle28);
            cellStyles1.Append(cellStyle29);
            cellStyles1.Append(cellStyle30);
            cellStyles1.Append(cellStyle31);
            cellStyles1.Append(cellStyle32);
            cellStyles1.Append(cellStyle33);
            cellStyles1.Append(cellStyle34);
            cellStyles1.Append(cellStyle35);
            cellStyles1.Append(cellStyle36);
            cellStyles1.Append(cellStyle37);
            cellStyles1.Append(cellStyle38);
            cellStyles1.Append(cellStyle39);
            cellStyles1.Append(cellStyle40);
            cellStyles1.Append(cellStyle41);
            cellStyles1.Append(cellStyle42);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />");

            stylesheetExtension2.Append(openXmlUnknownElement1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);





            workbookStylesPart.Stylesheet = stylesheet1;































            //Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            //stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            //stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            //Font font1 = new Font();
            //FontSize fontSize1 = new FontSize() { Val = 11D };
            //Color color1 = new Color() { Theme = (UInt32Value)1U };
            //FontName fontName1 = new FontName() { Val = "Calibri" };
            //FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            //FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            //font1.Append(fontSize1);
            //font1.Append(color1);
            //font1.Append(fontName1);
            //font1.Append(fontFamilyNumbering1);
            //font1.Append(fontScheme1);

            //fonts1.Append(font1);

            //Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            //Fill fill1 = new Fill();
            //PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            //fill1.Append(patternFill1);

            //Fill fill2 = new Fill();
            //PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            //fill2.Append(patternFill2);

            //fills1.Append(fill1);
            //fills1.Append(fill2);

            //Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            //Border border1 = new Border();
            //LeftBorder leftBorder1 = new LeftBorder();
            //RightBorder rightBorder1 = new RightBorder();
            //TopBorder topBorder1 = new TopBorder();
            //BottomBorder bottomBorder1 = new BottomBorder();
            //DiagonalBorder diagonalBorder1 = new DiagonalBorder();



            //border1.Append(leftBorder1);
            //border1.Append(rightBorder1);
            //border1.Append(topBorder1);
            //border1.Append(bottomBorder1);
            //border1.Append(diagonalBorder1);

            //Border border2 = new Border();

            //LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            //Color color2 = new Color() { Indexed = (UInt32Value)64U };

            //leftBorder2.Append(color2);

            //RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            //Color color3 = new Color() { Indexed = (UInt32Value)64U };

            //rightBorder2.Append(color3);

            //TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            //Color color4 = new Color() { Indexed = (UInt32Value)64U };

            //topBorder2.Append(color4);

            //BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            //Color color5 = new Color() { Indexed = (UInt32Value)64U };

            //bottomBorder2.Append(color5);
            //DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            //if (true)
            //{
            //    border2.Append(leftBorder2);
            //}

            //if (true)
            //{
            //    border2.Append(rightBorder2);
            //}
            //if (true)
            //{
            //    border2.Append(topBorder2);
            //}
            //if (true)
            //{
            //    border2.Append(bottomBorder2);
            //}
            //if (true)
            //{
            //    border2.Append(diagonalBorder2);
            //}






            //borders1.Append(border1);
            //borders1.Append(border2);

            //CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            //CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            //cellStyleFormats1.Append(cellFormat1);

            //CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
            //CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            //CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };

            //cellFormats1.Append(cellFormat2);
            //cellFormats1.Append(cellFormat3);

            //CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            //CellStyle cellStyle1 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            //cellStyles1.Append(cellStyle1);
            //DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            //TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

            //StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            //StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            //stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            //X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            //stylesheetExtension1.Append(slicerStyles1);

            //StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            //stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            //OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />");

            //stylesheetExtension2.Append(openXmlUnknownElement1);

            //stylesheetExtensionList1.Append(stylesheetExtension1);
            //stylesheetExtensionList1.Append(stylesheetExtension2);

            //stylesheet1.Append(fonts1);
            //stylesheet1.Append(fills1);
            //stylesheet1.Append(borders1);
            //stylesheet1.Append(cellStyleFormats1);
            //stylesheet1.Append(cellFormats1);
            //stylesheet1.Append(cellStyles1);
            //stylesheet1.Append(differentialFormats1);
            //stylesheet1.Append(tableStyles1);
            //stylesheet1.Append(stylesheetExtensionList1);

            //workbookStylesPart.Stylesheet = stylesheet1;
        }

        // Generates content of themePart1.
        private static void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Тема Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Стандартная" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Стандартная" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme2.Append(majorFont1);
            fontScheme2.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Стандартная" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme2);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

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
                int k = (int)d.BorderStyleId;

                Row row = GetRow(sheetData, i);
                Cell cell = new Cell() { CellReference = GetExcelColumnName(j) + i, StyleIndex = (UInt32Value)(UInt32)k };



                SetFormatedCellData(cell, d.Data);
                InsertCellIntoRow(cell, row);

            }

            Columns columns = new Columns();
            InsertColumnWidth(columns);

            SheetViews sheetViews = new SheetViews();
            sheetViews.Append(new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U });
            worksheet.Append(new SheetDimension() { Reference = "A1" });
            worksheet.Append(sheetViews);
            worksheet.Append(new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D });
            if (columnWidthArr.Count > 0)
                worksheet.Append(columns);
            worksheet.Append(sheetData);
            worksheet.Append(new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D });
            worksheetPart.Worksheet = worksheet;
            worksheet.Save();

        }

        private static void SetFormatedCellData(Cell cell, string data)
        {
            int number;
            decimal dec;

            if (data == null || data.Length == 0)
                return;


            CellValue cellValue = new CellValue();
            CellFormula cellFormula = new CellFormula();
            if (data.Substring(0, 1) == "=")
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cellFormula.Text = data.Remove(0, 1);
                cell.Append(cellFormula);
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

        //TODO 
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddColumnWidth(int columnIndex, double columnWidth)
        {

            columnWidthArr.Add(new ColumnWidth(columnIndex, columnWidth));

            return 0;
        }

        private static void InsertColumnWidth(Columns columns)
        {

            foreach (ColumnWidth d in columnWidthArr)
            {
                int c = (int)d.Column;
                double w = (double)d.Width;
                Column column = new Column()
                {
                    Min = (UInt32Value)(UInt32)c,
                    Max = (UInt32Value)(UInt32)c,
                    Width = w,
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
        public static int AddCellData(int rowIndex, int colIndex, string data, int borderStyleId)
        {

            cellsData.Add(new CellData(rowIndex, colIndex, data, borderStyleId));


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
                    //create your row here
                    newRow = GetNewRow(rowIndex);
                    sheetData.InsertBefore<Row>(newRow, current);
                    return newRow;
                }
            }

            //create your row here
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


        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static long ClearArray()
        {
            cellsData.Clear();
            rowHeightArr.Clear();
            columnWidthArr.Clear();

            return 0;
        }




    }
}
public class CellData
{
    int rowIndex;
    int colIndex;
    String data;
    int borderStyleId;


    public CellData(int rowIndex, int colIndex, String data, int borderStyleId)
    {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
        this.data = data;
        this.borderStyleId = borderStyleId;
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

    public String Data
    {
        get { return data; }
        set { data = value; }
    }

    public int BorderStyleId
    {
        get { return borderStyleId; }
        set { borderStyleId = value; }
    }



}
public class ColumnWidth
{
    int columnIndex;
    double columnWidth;

    public ColumnWidth(int columnIndex, double columnWidth)
    {
        this.columnIndex = columnIndex;
        this.columnWidth = columnWidth;
    }
    public int Column
    {
        get { return columnIndex; }
        set { columnIndex = value; }
    }
    public double Width
    {
        get { return columnWidth; }
        set { columnWidth = value; }
    }

}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace OpenXMLImportDLL
{
    public class ExcelImport
    {
        static private SortedList<int, double?> columnWidthArr = new SortedList<int, double?>();
        static private SortedList<int, double?> rowHeightArr = new SortedList<int, double?>();
        static private List<NPageSetup> nPageSetup = new List<NPageSetup>();
        static private List<string> mergeArr = new List<string>();
        static private List<CellData> cellsData = new List<CellData>();
        static private List<CellStyleFormat> cellStyleFormatList = new List<CellStyleFormat>();
        static private List<FontStyleFormat> fontStyleFormatList = new List<FontStyleFormat>();


        /// <summary>
        /// Convert Excel column number to Excel column name.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns>"Column name. (A,B,AF)"</returns>
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



        /// <summary>
        /// Convert Excel column name to Excel column number.
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns>Column number. (1,2,56)</returns>
        private static int GetExcelColumnNumber(string columnName)
        {
            string output = Regex.Replace(columnName, @"[\d-]", string.Empty);
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

        // Adds child parts and generates content of the specified part.
        /// <summary>
        /// Generate xslx file and after clear all arrays.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int GenerateExcel(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>("rId3");
                SharedStringTablePart sharedStringTablePart1 = workbookPart.AddNewPart<SharedStringTablePart>("rId4");

                Thread th1 = new Thread(GenerateWorksheetPartContent);
                th1.Start(worksheetPart as object);

                GenerateWorkbookPartContent(workbookPart);
                GenerateWorkbookStylesPartContent(workbookStylesPart);
                GenerateSharedStringTablePart1Content(sharedStringTablePart1);
                SetPackageProperties(document);

                th1.Join();
            }

            ClearArray();
            return 0;
        }

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
            }

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats();
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats = new CellFormats();
            CellFormat dCellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            cellFormats.Append(dCellFormat);

            foreach (CellStyleFormat f in cellStyleFormatList)
            {
                int treeadsw = 0;
                if (f.Treead)
                {
                    treeadsw = 4;
                }

                int lineStyle;
                switch (f.LineStyle)
                {
                    default: lineStyle = 0; break;
                    case 1: lineStyle = 1; break;
                    case 2: lineStyle = 2; break;
                    case 3: lineStyle = 3; break;
                    case 4: lineStyle = 4; break;
                    case 5: lineStyle = 5; break;
                }

                CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)(UInt32)treeadsw, FontId = (UInt32Value)(UInt32)f.FontIndex, FillId = (UInt32Value)0U, BorderId = (UInt32Value)(UInt32)lineStyle, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyNumberFormat = f.Treead };
                cellFormats.Append(cellFormat);

                VerticalAlignmentValues vav;
                HorizontalAlignmentValues hav;

                switch (f.HorizontalAlignment)
                {
                    default: hav = HorizontalAlignmentValues.Distributed; break;
                    case 1: hav = HorizontalAlignmentValues.Left; break;
                    case 2: hav = HorizontalAlignmentValues.Right; break;
                    case 3: hav = HorizontalAlignmentValues.Center; break;
                }

                switch (f.VerticalAlignment)
                {
                    default: vav = VerticalAlignmentValues.Distributed; break;
                    case 1: vav = VerticalAlignmentValues.Bottom; break;
                    case 2: vav = VerticalAlignmentValues.Center; break;
                    case 3: vav = VerticalAlignmentValues.Top; break;
                }

                Alignment alignment = new Alignment() { Horizontal = hav, Vertical = vav, WrapText = f.WrapText };
                cellFormat.Append(alignment);
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

        private static void GenerateWorksheetPartContent(object obj)
        {
            WorksheetPart worksheetPart = obj as WorksheetPart;
            Worksheet worksheet = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetData sheetData = new SheetData();
            PageSetup pageSetup = new PageSetup();
            OrientationValues orientation = OrientationValues.Default;
            bool showGreed = true;
            if (nPageSetup.Count > 0)
            {
                SheetProperties sheetProperties = new SheetProperties();
                PageSetupProperties pageSetupProperties = new PageSetupProperties() { FitToPage = true };
                sheetProperties.Append(pageSetupProperties);
                NPageSetup ps = nPageSetup[0];
                if (ps.LandscapeOrientation)
                    orientation = OrientationValues.Landscape;
                pageSetup = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = orientation, FitToWidth = (UInt32Value)(UInt32)ps.FitToPagesWide, FitToHeight = (UInt32Value)(UInt32)ps.FitToPagesTail, HorizontalDpi = (UInt32Value)300U, VerticalDpi = (UInt32Value)300U, Id = "rId1" };
                if (!ps.Grid)
                    showGreed = false;
                worksheet.Append(sheetProperties);
            }

            SheetViews sheetViews = new SheetViews();
            sheetViews.Append(new SheetView() { ShowGridLines = showGreed, TabSelected = true, WorkbookViewId = (UInt32Value)0U });
            worksheet.Append(new SheetDimension() { Reference = "A1" });
            worksheet.Append(sheetViews);

            int prCount = Environment.ProcessorCount;
            Thread[] threads = new Thread[prCount - 1];

            int part = cellsData.Count / prCount;
            int begin = 0, thrNum = 0;
            for (int i = 0; i < prCount; i++)
            {
                if (i == prCount - 1)
                {
                    part = cellsData.Count - begin;
                    WriteCellInTable(new Object[] { (Object)begin, (Object)part, (Object)sheetData }); //В текущем потоке
                    break;
                }
                threads[thrNum] = new Thread(WriteCellInTable);
                threads[thrNum].Start(new Object[] { (Object)begin, (Object)part, (Object)sheetData }); //Новый поток
                begin += part;
                thrNum++;
            }

            Columns columns = new Columns();
            InsertColumnWidth(columns);

            MergeCells mergeCells = new MergeCells();
            SetMergeCell(mergeCells);

            for (int i = 0; i < thrNum; i++)
                threads[i].Join();

            worksheet.Append(new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D });
            
            if (columnWidthArr.Count > 0)
                worksheet.Append(columns);
            worksheet.Append(sheetData);
            if (mergeArr.Count > 0)
                worksheet.Append(mergeCells);
            worksheetPart.Worksheet = worksheet;
            worksheet.Append(pageSetup);
            worksheet.Save();
        }


        private static void WriteCellInTable(Object obj)
        {
            Object[] param = (Object[])obj;
            int begin = (int)param[0];
            int part = (int)param[1];
            SheetData sheetData = (SheetData)param[2];
            Row previousRow = null;
            List<CellData> newList = cellsData.GetRange(begin, part);
            foreach (CellData currentCellData in newList)
            {
                int i = (int)currentCellData.Row;
                int j = (int)currentCellData.Column;
                int k = (int)currentCellData.Styleindex;

                if (previousRow == null || previousRow.RowIndex != i)
                {
                    previousRow = GetRow(sheetData, i);
                }
                string excelColName = GetExcelColumnName(j);
                Cell cell = new Cell() { CellReference = excelColName + i, StyleIndex = (UInt32Value)(UInt32)k };
                if (currentCellData.Data == null)
                    currentCellData.Data = "";
                SetFormatedCellData(cell, currentCellData.Data.ToString(), i, excelColName);
                InsertCellIntoRow(cell, previousRow);
            }
        }

        private static void SetFormatedCellData(Cell cell, string data, int i, string excelColName)
        {
            lock (cell)
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
                        string convertedFormula = TransformFormulaToA1Format(data, i, excelColName);
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

        /// <summary>
        /// Add merge cells. 
        /// </summary>
        /// <param name="mergedCellsId"></param>
        /// <returns></returns>
        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddMergeCell(string mergedCellsId)
        {
            mergeArr.Add(mergedCellsId);
            return 0;
        }

        [System.Reflection.Obfuscation(Feature = "DllExport")]
        public static int AddPageSetup(bool landscape, int fitToPageWide, int fitToPageTail, bool grid)
        {
            nPageSetup.Clear();
            nPageSetup.Add(new NPageSetup(landscape, fitToPageTail, fitToPageWide, grid));
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
            bool treead)
        {
            int ffi = SetFontFormat(bold, size, fontName, italic, underline);
            int cfi = SetCellFormat(wrapText, lineStyle, horizontalAlignment, verticalAlignment, treead, ffi);
            cellsData.Add(new CellData(rowIndex, colIndex, data, cfi));
            return 0;
        }


        private static Row GetRow(SheetData sheetData, int rowIndex)
        {
            lock (sheetData)
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
            lock (row)
            {
                foreach (Cell current in row.Elements<Cell>())
                {
                    int comp = current.CellReference.ToString().Length - cell.CellReference.ToString().Length;

                    if (comp == 0)
                        comp = current.CellReference.ToString().CompareTo(cell.CellReference.ToString());

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
            int fontIndex)
        {
            CellStyleFormat currentCell = new CellStyleFormat(wrapText, lineStyle, horizontalAlignment, verticalAlignment, treead, fontIndex);

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
            nPageSetup.Clear();
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
public class NPageSetup
{
    bool landscapeOrientation;
    int fitToPagesWide;
    int fitToPagesTail;
    bool grid;
    public NPageSetup(bool landscapeOrientation, int fitToPagesWide, int fitToPagesTail, bool grid)
    {
        this.landscapeOrientation = landscapeOrientation;
        this.fitToPagesWide = fitToPagesWide;
        this.fitToPagesTail = fitToPagesTail;
        this.grid = grid;

    }
    public bool LandscapeOrientation
    {
        get { return landscapeOrientation; }
        set { landscapeOrientation = value; }
    }
    public int FitToPagesWide
    {
        get { return fitToPagesWide; }
        set { fitToPagesWide = value; }
    }

    public int FitToPagesTail
    {
        get { return fitToPagesTail; }
        set { fitToPagesTail = value; }
    }

    public bool Grid
    {
        get { return grid; }
        set { grid = value; }
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
public class CellStyleFormat
{
    bool wrapText;
    int lineStyle;
    int horizontalAlignment;
    int verticalAlignment;
    bool treead;
    int fontIndex;
    public CellStyleFormat(
        bool wrapText,
        int lineStyle,
        int horizontalAlignment,
        int verticalAlignment,
        bool treead,
        int fontIndex)
    {
        this.wrapText = wrapText;
        this.lineStyle = lineStyle;
        this.horizontalAlignment = horizontalAlignment;
        this.verticalAlignment = verticalAlignment;
        this.treead = treead;
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
        return this.WrapText == cfs.wrapText && this.LineStyle == cfs.lineStyle && this.HorizontalAlignment == cfs.horizontalAlignment && this.VerticalAlignment == cfs.verticalAlignment && this.Treead == cfs.treead && this.fontIndex == cfs.fontIndex;
    }
    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}
public class FontStyleFormat
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
    public override int GetHashCode()
    {
        return base.GetHashCode();
    }
}

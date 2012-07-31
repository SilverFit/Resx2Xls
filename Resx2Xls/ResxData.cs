﻿namespace Resx2Xls
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Resources;
    using System.Runtime.InteropServices;
    using System.Text.RegularExpressions;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Convert resource data to and from .resx and .xlsx format
    /// </summary>
    public partial class ResxData
    {
        #region Excel rows
        private const int ExcelMetadataRow = 1;

        private const int ExcelHeaderRow = 2;

        private const int ExcelDataRow = 3;
        #endregion

        #region Excel columns
        /// <summary>
        /// Column index for key
        /// </summary>
        private const int ExcelFilesourceColumn = 1;

        /// <summary>
        /// Column index for key
        /// </summary>
        private const int ExcelKeyColumn = 1;

        /// <summary>
        /// Column index for value
        /// </summary>
        private const int ExcelValueColumn = 2;

        /// <summary>
        /// Column index for comment
        /// </summary>
        private const int ExcelCommentColumn = 3;

        /// <summary>
        /// Column index for first culture translation
        /// </summary>
        private const int ExcelCultureColumn = 4;
        #endregion

        /// <summary>
        /// List of keys that will be ignored from resx files
        /// </summary>
        private List<string> excludeList;

        /// <summary>
        /// Culture columns that will be added in the Excel document
        /// </summary>
        private List<CultureInfo> exportCultures;

        /// <summary>
        /// Import log of reading the resx files
        /// </summary>
        private string ReadResxReport = string.Empty;

        /// <summary>
        /// Create a ResxData instance from an xls file
        /// </summary>
        /// <param name="xlsFile">xls file to read</param>
        /// <returns>a newly instantized ResxData</returns>
        public static ResxData FromXls(string xlsFile)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(xlsFile, 0, false, 5, string.Empty, string.Empty, false, Excel.XlPlatform.xlWindows, string.Empty, true, false, 0, true, false, false);
            Excel.Sheets sheets = wb.Worksheets;

            ResxData rd = new ResxData();
            TranslationSourceRow currentResx = null;

            foreach (Excel.Worksheet sheet in sheets)
            {
                // Get filesource for current sheet
                var filesource = (sheet.Cells[ExcelMetadataRow, ExcelFilesourceColumn] as Excel.Range).Text.ToString();

                // Create a list of all cultures in the excel sheet
                List<string> cultures = new List<string>();
                int culturescolumn = ExcelCultureColumn;
                while (!String.IsNullOrEmpty((sheet.Cells[2, culturescolumn] as Excel.Range).Text.ToString()))
                {
                    cultures.Add((sheet.Cells[ExcelMetadataRow, culturescolumn] as Excel.Range).Text.ToString());
                    culturescolumn++;
                }

                // Iterate over all rows in the Excel sheet
                int row = ExcelDataRow;
                while (!String.IsNullOrEmpty((sheet.Cells[row, 1] as Excel.Range).Text.ToString()))
                {
                    if (currentResx == null || currentResx.FileSource != filesource)
                    {
                        currentResx = rd.TranslationSource.NewTranslationSourceRow();
                        currentResx.FileSource = filesource;
                        rd.TranslationSource.AddTranslationSourceRow(currentResx);
                    }

                    var resxKey = rd.PrimaryTranslation.NewPrimaryTranslationRow();
                    resxKey.Key = (sheet.Cells[row, ExcelKeyColumn] as Excel.Range).Text.ToString();
                    resxKey.Value = (sheet.Cells[row, ExcelValueColumn] as Excel.Range).Text.ToString();
                    resxKey.ResxRow = currentResx;
                    rd.PrimaryTranslation.AddPrimaryTranslationRow(resxKey);

                    // Iterate over all culture columns in the Excel sheet
                    for (int cultureindex = 0; cultureindex < cultures.Count; cultureindex++)
                    {
                        SecondaryTranslationRow lr = rd.SecondaryTranslation.NewSecondaryTranslationRow();
                        lr.Culture = cultures[cultureindex];
                        lr.Value = (sheet.Cells[row, ExcelCultureColumn + cultureindex] as Excel.Range).Text.ToString();
                        lr.PrimaryTranslationRow = resxKey;
                        rd.SecondaryTranslation.AddSecondaryTranslationRow(lr);
                    }

                    row++;
                }
            }

            rd.AcceptChanges();
            ExcelQuit(app, wb);
            return rd;
        }

        /// <summary>
        /// Read ResxData from .resx files
        /// </summary>
        /// <param name="path">root path of resx files</param>
        /// <param name="deepSearch">search subdirs</param>
        /// <param name="cultureList">list of cultures to translate</param>
        /// <param name="excludeList">list of keys to exclude</param>
        /// <param name="useFolderNamespacePrefix">use folder namespace prefix</param>
        /// <returns>a ResxData with all data</returns>
        public static ResxData FromResx(
            string path,
            bool deepSearch,
            bool purge,
            List<CultureInfo> cultureList,
            List<string> excludeList,
            bool useFolderNamespacePrefix)
        {
            ResxData rd = new ResxData();
            rd.exportCultures = cultureList;
            rd.excludeList = excludeList;

            var searchoptions = deepSearch ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(path, "*.resx", searchoptions);

            foreach (string f in files)
            {
                var resx = new Resx(f);
                if (!resx.IsCultureSpecific)
                {
                    rd.ReadResx(resx, path, purge, useFolderNamespacePrefix);
                }
            }

            if (!string.IsNullOrEmpty(rd.ReadResxReport))
            {
                Console.WriteLine(rd.ReadResxReport);
            }

            return rd;
        }

        /// <summary>
        /// Export this ResxData to an .xlsx file
        /// </summary>
        /// <param name="outputPath">Path to write .xlsx file to</param>
        public void ToXls(string outputPath, string screenshotPath)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Sheets sheets = wb.Worksheets;
            var cultures = this.exportCultures.Select(c => c.Name).ToList();
            int sheetIndex = 1;

            var firstSheet = app.ActiveSheet as Excel.Worksheet;

            // Iterate over all filesources that have keys assigned
            var filesources = this.PrimaryTranslation.Select(r => r.ResxRow.FileSource)
                                                     .Distinct();
            var filesourcesdict = new Dictionary<string, string>();
            foreach (var filesource in filesources)
            {
                var name = Regex.Replace(filesource, @"^.*\\", "");
                name = Regex.Replace(name, @"\.[^\.]*$", "");
                name = name.Substring(0, Math.Min(name.Length, 31));
                Debug.Assert(!filesourcesdict.ContainsKey(name), "Resource files with same name exist.");
                filesourcesdict.Add(name, filesource);
            }

            foreach (var filesource in filesourcesdict.OrderBy(kvp => kvp.Key))
            {
                // add sheet
                var sheet = sheets.Add(Before: sheets[sheetIndex]) as Excel.Worksheet;
                sheet.Name = filesource.Key;
                sheetIndex++;
                Trace.WriteLine("Created sheet " + filesource.Key);

                // add filesource metadata
                Excel.Range filesourcecell = sheet.Cells[ExcelMetadataRow, ExcelFilesourceColumn] as Excel.Range;
                filesourcecell.Value2 = filesource.Value;
                filesourcecell.Font.Italic = true;

                // add headers and culture metadata
                sheet.Cells[ExcelHeaderRow, ExcelKeyColumn] = "Key";
                sheet.Cells[ExcelHeaderRow, ExcelCommentColumn] = "Comment";
                sheet.Cells[ExcelHeaderRow, ExcelValueColumn] = "Value";
                int index = ExcelCultureColumn;
                foreach (var culture in this.exportCultures)
                {
                    sheet.Cells[ExcelHeaderRow, index] = culture.DisplayName;
                    sheet.Cells[ExcelMetadataRow, index] = culture.Name;
                    index++;
                }

                // make header bold and metadata italic
                var metadatarow = sheet.Rows[RowIndex: ExcelMetadataRow] as Excel.Range;
                metadatarow.Font.Italic = true;
                metadatarow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                metadatarow.Locked = true;
                metadatarow.Hidden = true;
                var headerrow = sheet.Rows[RowIndex: ExcelHeaderRow] as Excel.Range;
                headerrow.Font.Bold = true;
                headerrow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                headerrow.Locked = true;

                // set border
                var borders = headerrow.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom);
                borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                borders.Weight = Excel.XlBorderWeight.xlMedium;

                // add actual data
                int row = ExcelDataRow;
                foreach (var r in this.PrimaryTranslation.Where(r => r.ResxRow.FileSource == filesource.Value).OrderBy(r => r.Key))
                {
                    sheet.Cells[row, ExcelKeyColumn] = r.Key;
                    sheet.Cells[row, ExcelCommentColumn] = r.Comment;
                    sheet.Cells[row, ExcelValueColumn] = r.Value.Replace(@"\r\n", Environment.NewLine);
                    
                    SecondaryTranslationRow[] rows = r.GetResxLocalizedRows();

                    // Set background and unlock culture cells
                    for (int i = 0; i < cultures.Count; i++)
                    {
                        var cell = sheet.Cells[row, ExcelCultureColumn + i] as Excel.Range;
                        var color = r.Implicit ? System.Drawing.Color.Bisque : System.Drawing.Color.Yellow;
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
                        cell.Locked = false;
                        Marshal.ReleaseComObject(cell);
                    }

                    foreach (SecondaryTranslationRow lr in rows)
                    {
                        int col = cultures.IndexOf(lr.Culture);
                        if (col >= 0)
                        {
                            if (!string.IsNullOrEmpty(lr.Value))
                            {
                                var cell = sheet.Cells[row, ExcelCultureColumn + col] as Excel.Range;
                                var color = r.Implicit ? System.Drawing.Color.LightBlue : System.Drawing.Color.YellowGreen;
                                cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
                                sheet.Cells[row, col + ExcelCultureColumn] = lr.Value;
                                Marshal.ReleaseComObject(cell);
                            }
                        }
                    }

                    row++;
                }
                sheet.Cells.get_Range("A1", "Z1").EntireColumn.AutoFit();
                sheet.Cells.get_Range("A1", "Z1").EntireColumn.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                // Insert screenshots
                List<string> validExtensions = new List<string>();
                validExtensions.Add(".png");
                validExtensions.Add(".jpg");
                validExtensions.Add(".jpeg");

                var screenshotDir = Path.Combine(screenshotPath, filesource.Key);
                if (Directory.Exists(screenshotDir))
                {
                    var lastcell = sheet.Cells[row, 1] as Excel.Range;
                    var lastCellTopPoints = (Double)(lastcell.Top);
                    var offset = (float)lastCellTopPoints + 20;
                    Marshal.ReleaseComObject(lastcell);
                    var screenshotFiles = Directory.GetFiles(screenshotDir).Where(f => validExtensions.Contains(Path.GetExtension(f)));
                    foreach (var file in screenshotFiles)
                    {
                        var size = System.Drawing.Image.FromFile(file).Size;
                        sheet.Shapes.AddPicture(file,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoCTrue,
                            10, offset, size.Width / 2, size.Height / 2);
                        offset += size.Height / 2 + 10;
                    }
                }
 
                // Set width of value column
                var valuecolumn = sheet.Columns.get_Item(ExcelValueColumn) as Excel.Range;
                valuecolumn.ColumnWidth = 80;

                // Set width of translated columns
                for (int i = 0; i < cultures.Count; i++)
                {
                    var column = sheet.Columns.get_Item(i + ExcelCultureColumn) as Excel.Range;
                    column.ColumnWidth = 60;
                    Marshal.ReleaseComObject(column);
                }

                // hide key column
                if (Properties.Settings.Default.HideKeys)
                {
                    var column = sheet.Columns.get_Item(ExcelKeyColumn) as Excel.Range;
                    column.Hidden = true;
                    Marshal.ReleaseComObject(column);
                }

                // hide comment column
                if (Properties.Settings.Default.HideComments)
                {
                    var column = sheet.Columns.get_Item(ExcelCommentColumn) as Excel.Range;
                    //column.Hidden = true;
                    column.ColumnWidth = 0;
                    Marshal.ReleaseComObject(column);
                }

                sheet.Protect(Password: "", Contents: true, AllowFormattingColumns: true);

                Marshal.ReleaseComObject(metadatarow);
                Marshal.ReleaseComObject(headerrow);
                Marshal.ReleaseComObject(borders);
                Marshal.ReleaseComObject(valuecolumn);
                Marshal.ReleaseComObject(sheet);
            }

            // Remove Sheet1 that is added by default
            firstSheet.Delete();

            // Save the Workbook and force overwriting by rename trick
            string tmpFile = Path.GetTempFileName();
            File.Delete(tmpFile);
            wb.SaveAs(Filename: tmpFile, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange);
            File.Delete(outputPath);

            ExcelQuit(app, wb);

            // Move file otherwise handle is still used by excel
            File.Move(tmpFile, outputPath);

            if (!string.IsNullOrEmpty(this.ReadResxReport))
            {
                File.WriteAllText(outputPath + ".log", this.ReadResxReport);
            }
        }

        /// <summary>
        /// Export to resx files
        /// </summary>
        /// <param name="path">path to place resx files</param>
        public void ToResx(string path)
        {
            var cultures = this.SecondaryTranslation.Select(rl => rl.Culture)
                                                    .Distinct();

            foreach (var culture in cultures)
            {
                string pathCulture = Path.Combine(path, culture);

                if (!System.IO.Directory.Exists(pathCulture))
                {
                    System.IO.Directory.CreateDirectory(pathCulture);
                }

                foreach (var resx in this.TranslationSource)
                {
                    var fullPath = Path.Combine(pathCulture, resx.FileSource);
                    var directoryName = Path.GetDirectoryName(fullPath);
                    var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fullPath);
                    var fileName = fileNameWithoutExtension + "." + culture + ".resx";

                    Directory.CreateDirectory(directoryName);
                    ResXResourceWriter rw = new ResXResourceWriter(Path.Combine(directoryName, fileName));

                    foreach (var entry in this.SecondaryTranslation.Where(r => r.PrimaryTranslationRow.TranslationSource == resx.Id)
                                                                   .Where(r => r.Culture == culture)
                                                                   .Where(r => !string.IsNullOrEmpty(r.Value)))
                    {
                        var value = entry.Value;
                        value = value.Replace("\\r", "\r");
                        value = value.Replace("\\n", "\n");
                        rw.AddResource(new ResXDataNode(entry.PrimaryTranslationRow.Key, value));
                    }

                    Console.WriteLine("Wrote localized resx {0}", fileName);
                    rw.Close();
                }
            }
        }

        /// <summary>
        /// Close Excel workbook, quit excel and wait until finished
        /// </summary>
        /// <param name="app">Excel handle to quit</param>
        /// <param name="wb">Excel workbook to close</param>
        private static void ExcelQuit(Excel.Application app, Excel.Workbook wb)
        {
            wb.Close(SaveChanges: false);
            app.Quit();
            while (Marshal.ReleaseComObject(app) != 0)
            {
            }

            app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Escape the \n and \r character
        /// </summary>
        /// <param name="value">string to escape</param>
        /// <returns>escaped string</returns>
        private static string EscapeNewline(string value)
        {
            value = value.Replace("\r", "\\r");
            value = value.Replace("\n", "\\n");
            return value;
        }

        /// <summary>
        /// Read a specific .resx file
        /// </summary>
        /// <param name="primaryResx">resource file to import</param>
        /// <param name="projectRoot">root of project used to calculate relative path</param>
        /// <param name="useFolderNamespacePrefix">use folder namespace prefix</param>
        private void ReadResx(
            Resx primaryResx,
            string projectRoot,
            bool purge,
            bool useFolderNamespacePrefix)
        {
            // Create resx reader for primary language
            var primaryEntries = primaryResx.Read()
                                      .Where(k => this.ValidateKey(k.Key))
                                      .Where(k => !string.IsNullOrEmpty(k.Value));

            if (primaryEntries.Count() > 0)
            {
                // Create translation source entry for resx file
                TranslationSourceRow resxrow = this.TranslationSource.NewTranslationSourceRow();
                resxrow.FileSource = primaryResx.GetRelativePath(projectRoot);
                this.TranslationSource.AddTranslationSourceRow(resxrow);
                
                // Iterate over all entries in resource file fileName
                foreach (ResxItem resxitem in primaryEntries)
                {
                    var resxkey = this.AddResxKey(resxrow, resxitem.Key, resxitem.Value, resxitem.Comment);
                }

                // Create resx readers for all requested cultures
                var secondaryResxs = new Dictionary<string, IEnumerable<ResxItem>>();
                foreach (CultureInfo culture in exportCultures)
                {
                    var path = string.Format("{0}.{1}.resx", primaryResx.PathWithoutExtension, culture.Name);
                    if (File.Exists(path))
                    {
                        var resx = new Resx(path);
                        var entries = resx.Read()
                                        .Where(k => this.ValidateKey(k.Key))
                                        .Where(k => !string.IsNullOrEmpty(k.Value));
                        
                        foreach (ResxItem resxitem in entries)
                        {
                            var row = this.PrimaryTranslation.Where(p => p.TranslationSource == resxrow.Id)
                                                             .FirstOrDefault(r => r.Key == resxitem.Key);

                            if (row == null && !purge)
                            {
                                row = this.AddResxKey(resxrow, resxitem.Key, "", "does not exist in primary language");
                                row.Implicit = true;
                            }

                            if (row != null)
                            {
                                this.AddSecondaryKey(row, culture.Name, resxitem.Value);
                            }
                        }
                    }
                }

                // Check for missing keys when base language has keys
                var untranslated = this.PrimaryTranslation.Where(p => p.ResxRow == resxrow)
                                                          .Where(p => p.GetResxLocalizedRows().Count() == 0);
                if (untranslated.Count() > 0)
                {
                    string missingReport = string.Format(
                        "Missing translations from {0}:{1}{2}",
                        primaryResx.GetRelativePath(projectRoot),
                        Environment.NewLine, 
                        string.Join(", ", untranslated.Select(u => u.Key).ToArray()));
                    this.ReadResxReport += missingReport + Environment.NewLine + Environment.NewLine;
                }
            }
        }

        /// <summary>
        /// Validates whether a key is valid for translation
        /// </summary>
        /// <param name="key">translation key</param>
        /// <returns>true when key should be translated</returns>
        private bool ValidateKey(string key)
        {
            return !excludeList.Any(e => key.EndsWith(e));
        }

        /// <summary>
        /// Adds values for the primary language
        /// </summary>
        private PrimaryTranslationRow AddResxKey(TranslationSourceRow row, string key, string value, string comment)
        {
            PrimaryTranslationRow primary = this.PrimaryTranslation.NewPrimaryTranslationRow();
            primary.Key = key;
            primary.Value = value;
            primary.Comment = comment;
            primary.ResxRow = row;
            this.PrimaryTranslation.AddPrimaryTranslationRow(primary);
            return primary;
        }

        /// <summary>
        /// Adds values for a secondary language
        /// </summary>
        /// <param name="culture">culture of the secondary language</param>
        private SecondaryTranslationRow AddSecondaryKey(PrimaryTranslationRow primary, string culture, string value)
        {
            SecondaryTranslationRow secondary = this.SecondaryTranslation.NewSecondaryTranslationRow();
            secondary.Culture = culture;
            secondary.PrimaryTranslationRow = primary;
            secondary.Value = value;
            this.SecondaryTranslation.AddSecondaryTranslationRow(secondary);
            return secondary;
        }
    }
}
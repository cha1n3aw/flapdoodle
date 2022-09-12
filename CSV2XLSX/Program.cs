using Microsoft.VisualBasic.FileIO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Globalization;
using OfficeOpenXml.Drawing.Chart;
using System.CommandLine;
using System.CommandLine.DragonFruit;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using OfficeOpenXml.Drawing.Chart.Style;

namespace CSV2XLSX
{
    internal class Program
    {
        /// <summary>
        /// A simple .csv to .xlsx parser with some chart magic
        /// </summary>
        /// <param name="i">Input directory (obligatory)</param>
        /// <param name="f">Filenames array (obligatory, or specify -a)</param>
        /// <param name="a">Select all files in an input directory (obligatory, or specify -f)</param>
        /// <param name="o">Output directory (optional)</param>
        /// <param name="dc">Disable chart creation (optional)</param>
        /// <param name="aw">Enable column auto-width (optional)</param>
        /// <param name="ro">Rows of data offset (optional)</param>
        /// <param name="fo">Columns of data offset (optional)</param>
        /// <param name="fa">Amount of data columns (optional)</param>
        /// <param name="fn">Column name for chart Y-axis (optional)</param>
        /// <param name="ds">Delimiter string for .csv file (optional)</param>
        /// <param name="wn">Worksheet name in .xlsx file (optional)</param>
        /// <param name="ct">Custom chart title (optional)</param>
        /// <param name="ch">Custom chart height (optional)</param>
        /// <param name="cw">Custom chart width (optional)</param>
        static void Main(
            string? i = null,
            string[]? f = null,
            bool a = false,
            string? o = null,
            bool dc = false,
            bool aw = false,
            int ro = 0,
            int fo = 0,
            int fa = 0,
            string fn = "ms",
            string ds = ";",
            string wn = "Worksheet 1",
            string ct = "Parameters chart",
            int ch = 500,
            int cw = 1000)
        {
            try
            {
                if (i == null || i.Length == 0)
                {
                    Console.WriteLine("Error: no -i specified");
                    return;
                }
                else if (!Path.Exists(i))
                {
                    Console.WriteLine($"Error: -i '{i}' doesn't exist");
                    return;
                }
                if (!a)
                {
                    if (f == null || f.Length == 0)
                    {
                        Console.WriteLine("Error: no -f specified");
                        return;
                    }
                    else 
                        foreach (string name in f)
                        if (!name.Contains(".csv"))
                        {
                            Console.WriteLine($"Error: -f '{name}' has wrong extension");
                            return;
                        }
                        else if (!File.Exists(Path.Combine(i, name)))
                        {
                            Console.WriteLine($"Error: -f '{name}' doesn't exist");
                            return;
                        }
                }
                else
                {
                    f = Directory.GetFiles(i, "*.csv");
                    for (int y = 0; y < f.Length; y++) f[y] = Path.GetFileName(f[y]);
                }
                if (o == null || o == string.Empty) o = i;
                else if (!Path.Exists(o))
                {
                    Console.WriteLine($"Error: -o '{o}' doesn't exist");
                    return;
                }
                ro++;
                int yAxisFieldColumn = 0;
                foreach (string csvFile in f)
                {
                    string csvFilePath = String.Empty;
                    string xlsxFilePath = String.Empty;
                    List<string[]> csvRows = new();
                    try
                    {
                        csvFilePath = Path.Combine(i, csvFile);
                        xlsxFilePath = Path.Combine(o, csvFile.Replace(".csv", ".xlsx"));
                        using (TextFieldParser parser = new(csvFilePath))
                        {
                            parser.TextFieldType = FieldType.Delimited;
                            parser.SetDelimiters(ds);
                            int headerRow = 0;
                            while (!parser.EndOfData)
                            {
                                string[]? fields = parser.ReadFields();
                                if (fields != null && fields.Length > 0)
                                {
                                    if (headerRow == 0 && fields.Contains(fn))
                                    {
                                        if (fn != null && fn != string.Empty) yAxisFieldColumn = fields.ToList().IndexOf(fn) + 1;
                                        headerRow = csvRows.Count;
                                        fa = fields.Length - yAxisFieldColumn;
                                    }
                                    csvRows.Add(fields);
                                }
                            }
                            csvRows.RemoveRange(0, headerRow);
                        }
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"Exception thrown while processing a .csv file: {e.Message}, {e.StackTrace}");
                    }
                    try
                    {
                        if (File.Exists(xlsxFilePath)) File.Delete(xlsxFilePath);
                        FileInfo workbookFileInfo = new(xlsxFilePath);
                        using ExcelPackage excelPackage = new(workbookFileInfo);
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(wn);

                        for (int z = 0; z < csvRows.Count; z++)
                        {
                            for (int y = 0; y < csvRows[z].Length; y++)
                            {
                                csvRows[z][y] = Regex.Replace(csvRows[z][y], @"[^\u0000-\u007F]+", string.Empty);                       //delete all non-utf8 chars
                                if (z == 0) worksheet.Cells[z + 1, y + 1].Value = csvRows[z][y];
                                else
                                {
                                    if (Regex.IsMatch(csvRows[z][y], "^[\\+\\-]{0,1}\\d+[\\.\\,]\\d+$"))                                //double
                                        worksheet.Cells[z + 1, y + 1].Value = Double.Parse(csvRows[z][y], CultureInfo.InvariantCulture);
                                    else if (Regex.IsMatch(csvRows[z][y], "^[\\+\\-]{0,1}\\d+$"))                                       //int
                                        worksheet.Cells[z + 1, y + 1].Value = Int32.Parse(csvRows[z][y]);
                                    else if (Regex.IsMatch(csvRows[z][y], "^(?:(?:31(\\/|-|\\.)(?:0?[13578]|1[02]))\\1|(?:(?:29|30)(\\/|-|\\.)(?:0?[13-9]|1[0-2])\\2))(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$|^(?:29(\\/|-|\\.)0?2\\3(?:(?:(?:1[6-9]|[2-9]\\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\\d|2[0-8])(\\/|-|\\.)(?:(?:0?[1-9])|(?:1[0-2]))\\4(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$"))                                            //date
                                    {
                                        DateTime dt = DateTime.Parse(csvRows[z][y], CultureInfo.InvariantCulture);
                                        worksheet.Cells[z + 1, y + 1].Style.Numberformat.Format = "dd.mm.yyyy";
                                        worksheet.Cells[z + 1, y + 1].Formula = $"=DATE({dt.Year},{dt.Month},{dt.Day})";
                                    }
                                    else if (Regex.IsMatch(csvRows[z][y], "^(?:(?:([01]?\\d|2[0-3]):)?([0-5]?\\d):)?([0-5]?\\d)$"))     //time
                                    {
                                        DateTime dt = DateTime.ParseExact(csvRows[z][y], "HH:mm:ss", CultureInfo.InvariantCulture);
                                        worksheet.Cells[z + 1, y + 1].Style.Numberformat.Format = "hh:mm:ss";
                                        worksheet.Cells[z + 1, y + 1].Formula = $"=TIME({dt.Hour},{dt.Minute},{dt.Second})";
                                    }
                                }
                            }
                        }
                        if (!dc)
                        {
                            ExcelChart chart = worksheet.Drawings.AddChart(ct, eChartType.XYScatterLinesNoMarkers);
                            chart.Legend.Position = eLegendPosition.Bottom;
                            chart.Title.Text = ct;
                            chart.SetPosition(1, 0, yAxisFieldColumn + fa + 1, 0);
                            chart.SetSize(cw, ch);
                            chart.XAxis.MaxValue = Convert.ToDouble(worksheet.Cells[csvRows.Count, yAxisFieldColumn].Value);
                            chart.XAxis.MinValue = Convert.ToDouble(worksheet.Cells[ro + 1, yAxisFieldColumn].Value);
                            for (int y = yAxisFieldColumn + fo + 2; y <= yAxisFieldColumn + fa; y++)
                            {
                                ExcelChartSerie serie = chart.PlotArea.ChartTypes.Add(eChartType.XYScatterLinesNoMarkers).Series.Add(worksheet.Cells[ro + 1, y, csvRows.Count, y], worksheet.Cells[ro + 1, yAxisFieldColumn, csvRows.Count, yAxisFieldColumn]);
                                serie.Header = worksheet.Cells[ro, y].Value.ToString();  
                            }
                        }
                        if (aw) worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                        worksheet.View.FreezePanes(2, 1);
                        excelPackage.Save();
                        Console.WriteLine($"Success, {csvFilePath} saved as {xlsxFilePath}");
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine($"Exception thrown while processing an .xlsx file: {e.Message}, {e.StackTrace}");
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine($"Unhandled exception thrown: {e.Message}, {e.StackTrace}");
            }
        }
    }
}
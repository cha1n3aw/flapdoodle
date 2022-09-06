using Microsoft.VisualBasic.FileIO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.Globalization;
using OfficeOpenXml.Drawing.Chart;
using System.CommandLine;
using System.CommandLine.DragonFruit;

namespace CSV2XLSX
{
    internal class Program
    {
        /// <summary>
        /// A simple .csv to .xlsx parser with some chart magic
        /// </summary>
        /// <param name="inPath">Input directory</param>
        /// <param name="filename">Filenames array</param>
        /// <param name="all">Select all files in an input directory</param>
        /// <param name="outPath">Output directory</param>
        /// <param name="dataRowsOffset">Rows of data offset</param>
        /// <param name="dataFieldsOffset">Columns of data offset</param>
        /// <param name="dataFieldsAmount">Amount of data columns</param>
        /// <param name="yAxisFieldName">Column name for chart Y-axis</param>
        /// <param name="delimiterString">Delimiter string for .csv file</param>
        /// <param name="worksheetName">Worksheet name in .xlsx file</param>
        /// <param name="chartTitle">Chart title</param>
        static void Main(
            string? inPath = null,
            string[]? filename = null,
            bool all = false,
            string? outPath = null,
            int dataRowsOffset = 0,
            int dataFieldsOffset = 0,
            int dataFieldsAmount = 0,
            string yAxisFieldName = "ms",
            string delimiterString = ";",
            string worksheetName = "Worksheet 1",
            string chartTitle = "Parameters chart")
        {
            if ((filename == null || filename.Length == 0) && !all) throw new Exception("Specify --filenames OR --all");
            if (inPath == null || inPath.Length == 0) throw new Exception("Specify --in-path");
            if (outPath == null || outPath == string.Empty) outPath = inPath;
            if (all)
            {
                filename = Directory.GetFiles(inPath, "*.csv");
                for (int i = 0; i < filename.Length; i++) filename[i] = Path.GetFileName(filename[i]);
            }
            dataRowsOffset++;
            foreach (string csvFile in filename)
            {
                string csvFilePath = Path.Combine(inPath, csvFile);
                string xlsxFilePath = Path.Combine(outPath, csvFile.Replace(".csv", ".xlsx"));
                int yAxisFieldColumn = 0;
                List<string[]> csvRows = new();
                using (TextFieldParser parser = new(csvFilePath))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(delimiterString);
                    int headerRow = 0;
                    while (!parser.EndOfData)
                    {
                        string[]? fields = parser.ReadFields();
                        if (fields != null && fields.Length > 0)
                        {
                            if (headerRow == 0 && fields.Contains(yAxisFieldName))
                            {
                                if (yAxisFieldColumn == 0 && yAxisFieldName != null && yAxisFieldName != string.Empty) yAxisFieldColumn = fields.ToList().IndexOf(yAxisFieldName) + 1;
                                headerRow = csvRows.Count;
                                if (dataFieldsAmount == 0) dataFieldsAmount = fields.Length - yAxisFieldColumn;
                            }
                            csvRows.Add(fields);
                        }
                    }
                    csvRows.RemoveRange(0, headerRow);
                }

                if (File.Exists(xlsxFilePath)) File.Delete(xlsxFilePath);
                FileInfo workbookFileInfo = new(xlsxFilePath);
                using ExcelPackage excelPackage = new(workbookFileInfo);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(worksheetName);

                for (int i = 0; i < csvRows.Count; i++)
                {
                    for (int y = 0; y < csvRows[i].Length; y++)
                    {
                        csvRows[i][y] = Regex.Replace(csvRows[i][y], @"[^\u0000-\u007F]+", string.Empty);                       //delete all non-utf8 chars
                        if (i == 0) worksheet.Cells[i + 1, y + 1].Value = csvRows[i][y];
                        else
                        {
                            if (Regex.IsMatch(csvRows[i][y], "^[\\+\\-]{0,1}\\d+[\\.\\,]\\d+$"))                                //double
                                worksheet.Cells[i + 1, y + 1].Value = Double.Parse(csvRows[i][y], CultureInfo.InvariantCulture);
                            else if (Regex.IsMatch(csvRows[i][y], "^[\\+\\-]{0,1}\\d+$"))                                       //int
                                worksheet.Cells[i + 1, y + 1].Value = Int32.Parse(csvRows[i][y]);
                            else if (Regex.IsMatch(csvRows[i][y], "^(?:(?:31(\\/|-|\\.)(?:0?[13578]|1[02]))\\1|(?:(?:29|30)(\\/|-|\\.)(?:0?[13-9]|1[0-2])\\2))(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$|^(?:29(\\/|-|\\.)0?2\\3(?:(?:(?:1[6-9]|[2-9]\\d)?(?:0[48]|[2468][048]|[13579][26])|(?:(?:16|[2468][048]|[3579][26])00))))$|^(?:0?[1-9]|1\\d|2[0-8])(\\/|-|\\.)(?:(?:0?[1-9])|(?:1[0-2]))\\4(?:(?:1[6-9]|[2-9]\\d)?\\d{2})$"))                                            //date
                            {
                                DateTime dt = DateTime.Parse(csvRows[i][y], CultureInfo.InvariantCulture);
                                worksheet.Cells[i + 1, y + 1].Style.Numberformat.Format = "dd.mm.yyyy";
                                worksheet.Cells[i + 1, y + 1].Formula = $"=DATE({dt.Year},{dt.Month},{dt.Day})";
                            }
                            else if (Regex.IsMatch(csvRows[i][y], "^(?:(?:([01]?\\d|2[0-3]):)?([0-5]?\\d):)?([0-5]?\\d)$"))     //time
                            {
                                DateTime dt = DateTime.ParseExact(csvRows[i][y], "HH:mm:ss", CultureInfo.InvariantCulture);
                                worksheet.Cells[i + 1, y + 1].Style.Numberformat.Format = "hh:mm:ss";
                                worksheet.Cells[i + 1, y + 1].Formula = $"=TIME({dt.Hour},{dt.Minute},{dt.Second})";
                            }
                        }
                    }
                }

                ExcelChart chart1 = worksheet.Drawings.AddChart("Engine parameters", eChartType.XYScatterLinesNoMarkers);
                chart1.Title.Text = chartTitle;
                chart1.SetPosition(1, 0, yAxisFieldColumn + dataFieldsAmount + 1, 0);
                chart1.SetSize(1500, 1000);
                ExcelChartSerie serie1 = chart1.Series.Add(worksheet.Cells[dataRowsOffset + 1, yAxisFieldColumn + 1 + dataFieldsOffset, csvRows.Count, yAxisFieldColumn + dataFieldsOffset + 1], worksheet.Cells[dataRowsOffset + 1, yAxisFieldColumn, csvRows.Count, yAxisFieldColumn]);
                serie1.Header = worksheet.Cells[dataRowsOffset, yAxisFieldColumn + dataFieldsOffset + 1].Value.ToString();
                for (int i = yAxisFieldColumn + dataFieldsOffset + 2; i <= yAxisFieldColumn + dataFieldsAmount; i++)
                {
                    ExcelChart chart2 = chart1.PlotArea.ChartTypes.Add(eChartType.XYScatterLinesNoMarkers);
                    ExcelChartSerie serie2 = chart2.Series.Add(worksheet.Cells[dataRowsOffset + 1, i, csvRows.Count, i], worksheet.Cells[dataRowsOffset + 1, yAxisFieldColumn, csvRows.Count, yAxisFieldColumn]);
                    serie2.Header = worksheet.Cells[dataRowsOffset, i].Value.ToString();
                }
                excelPackage.Save();
                Console.WriteLine($"Success, {csvFilePath} saved as {xlsxFilePath}");
            }
        }
    }
}
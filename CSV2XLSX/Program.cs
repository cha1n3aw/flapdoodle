using Microsoft.VisualBasic.FileIO;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using Spire.Xls;
using OfficeOpenXml;

namespace ConsoleApp2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (TextFieldParser parser = new TextFieldParser(@"C:\Users\ryabuhin_ia\repos\flapdoodle\before_edit.csv"))
            {
                string csvFile = @"C:\Users\ryabuhin_ia\repos\flapdoodle\afetr_edit_228.csv";
                string xlsFile = @"C:\Users\ryabuhin_ia\repos\flapdoodle\afetr_edit_228.xlsx";

                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                int mainBody = 0;
                List<string[]> rows = new();
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    rows.Add(fields);
                    if (mainBody == 0)
                        for (int i = 0; i < fields.Length; i++)
                            if (fields[i].Contains("Date") || fields[i].Contains("Time") || fields[i].Contains("ms"))
                                mainBody = rows.Count() - 1;
                }
                rows.RemoveRange(0, mainBody);
                StringBuilder csv = new StringBuilder();
                for (int i = 0; i < rows.Count; i++)
                {
                    string line = string.Empty;
                    for (int y = 0; y < rows[i].Length; y++)
                    {
                        rows[i][y] = Regex.Replace(rows[i][y], @"[^\u0000-\u007F]+", string.Empty);
                        Console.WriteLine(rows[i][y]);
                        line += $"{rows[i][y]},";
                    }
                    csv.AppendLine(line.Trim(','));
                }
                File.WriteAllText(csvFile, csv.ToString());

                if (File.Exists(xlsFile)) File.Delete(xlsFile);

                Workbook workbook = new();
                workbook.LoadFromFile(csvFile, ",", 1, 1);
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Parsed data";
                workbook.SaveToFile(xlsFile, ExcelVersion.Version2010);

                FileInfo workbookFileInfo = new(xlsFile);
                using var excelPackage = new ExcelPackage(workbookFileInfo);
                excelPackage.Workbook.Worksheets.Delete(excelPackage.Workbook.Worksheets.SingleOrDefault(x => x.Name.Contains("Evaluation Warning")));
                excelPackage.Save();
            }
        }
    }
}
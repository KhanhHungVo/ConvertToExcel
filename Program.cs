using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ConvertToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = System.Configuration.ConfigurationManager.AppSettings["SourceDataPath"];
            
            
            DirectoryInfo d = new DirectoryInfo(path);//Assuming Test is your Folder
            string excelFileName = $"{path}test.xlsx";
            if (File.Exists(excelFileName))
            {
                File.Delete(excelFileName);
            }
            FileInfo[] Files = d.GetFiles("*.csv"); //Getting csv files
            

            foreach (FileInfo file in Files)
            {
                string worksheetsName = $"{file.Name.Replace(".csv", "")}";
                string csvFileName = $"{path}{file.Name}";
                var lines = File.ReadAllLines(csvFileName);
                
                    

                bool firstRowIsHeader = true;

                var format = new ExcelTextFormat();
                format.Delimiter = '\t';

                if (lines.Any(arg => arg.Contains("record(s) selected")))
                {
                    var splitRegex = @"[ ]+(?!((\d{2}):(\d{2})))";
                    var newContent = lines.Where(arg => !string.IsNullOrWhiteSpace(arg) && !arg.Contains("record(s) selected")
                    && !Regex.IsMatch(arg, @"[-]+"));
                    foreach(var row in newContent)
                    {
                        Regex.Replace(row, splitRegex, "\t");
                    }
                   // newContent.First() = Regex.Replace(newContent.First(), @" +", "\t");
                    File.WriteAllLines(csvFileName, newContent);
                }
                //format.EOL = @"\s+";              // DEFAULT IS "\r\n";
                // format.TextQualifier = '"';

                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
                {
                    if (package.Workbook.Worksheets.Any(sheet => sheet.Name == worksheetsName))
                    {
                        package.Workbook.Worksheets.Delete(worksheetsName);
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                    worksheet.Cells["A1"].LoadFromText(new FileInfo(csvFileName), format, OfficeOpenXml.Table.TableStyles.Medium27, firstRowIsHeader);
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    for (int row = start.Row; row <= end.Row; row++)
                    { // Row by row...
                        for (int col = start.Column; col <= end.Column; col++)
                        { // ... Cell by cell...
                           worksheet.Cells[row, col].Value = worksheet.Cells[row, col].Text.Replace("\"",""); // This got me the actual value I needed.
                        }
                    }

                    package.Save();
                }

            }

            Console.WriteLine("Finished!");
            Console.ReadLine();
        }
    }
}

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

                bool firstRowIsHeader = true;

                var format = new ExcelTextFormat();
                format.Delimiter = '\t';


                var text = File.ReadAllText(csvFileName);
               

                if (text.Contains("record(s) selected") || text.Contains("0 rows affected"))
                {
                    // Replace seperator by tab
                    var splitRegex = @"[ ]+(?!((\d{2}):(\d{2})))";
                    text = Regex.Replace(text, splitRegex, "\t");
                    File.WriteAllText(csvFileName, text);

                    var lines = File.ReadAllLines(csvFileName);

                    var newContent = lines.Where(arg => !string.IsNullOrWhiteSpace(arg) && !arg.Contains("record(s)") && !arg.Contains("rows")
                    && !Regex.IsMatch(arg, @"[-][-]+"));
                   
                   // newContent.First() = Regex.Replace(newContent.First(), @" +", "\t");
                    File.WriteAllLines(csvFileName, newContent);
                }


                //format.EOL = @"\s+";              // DEFAULT IS "\r\n";
                // format.TextQualifier = '"';


                if (file.Name.Contains("PremiumAmountsAndSunriseFields"))
                {
                    format.Delimiter = ',';
                }
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFileName)))
                {
                    if (package.Workbook.Worksheets.Any(sheet => sheet.Name == worksheetsName))
                    {
                        package.Workbook.Worksheets.Delete(worksheetsName);
                    }
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetsName);
                    worksheet.Cells["A1"].LoadFromText(
                        new FileInfo(csvFileName)
                        , format
                       // , OfficeOpenXml.Table.TableStyles.Medium27
                        , OfficeOpenXml.Table.TableStyles.None
                        , firstRowIsHeader);
                    var start = worksheet.Dimension.Start;
                    var end = worksheet.Dimension.End;
                    for (int row = start.Row; row <= end.Row; row++)
                    { // Row by row...

                        bool endRowEmpty = true;
                        for (int col = start.Column; col <= end.Column; col++)
                        { // ... Cell by cell...
                           if(row == start.Column && worksheet.Cells[row, col].Text.Contains("Column"))
                            {
                                worksheet.DeleteColumn(col);
                                break;
                            }
                            if (row == end.Row && worksheet.Cells[row, col].Value != null)
                            {
                                endRowEmpty = false;
                            }
                            worksheet.Cells[row, col].Value = worksheet.Cells[row, col].Text.Replace("\"","").Replace(" #",""); // This got me the actual value I needed.
                        }
                        if(row == end.Row && endRowEmpty)
                        {
                            worksheet.DeleteRow(row);
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

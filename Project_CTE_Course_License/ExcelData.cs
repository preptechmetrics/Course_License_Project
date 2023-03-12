using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace Project_CTE_Course_License
{
    public class ExcelData
    {
        public Dictionary<string, List<Dictionary<string, string>>> GetDictionary(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                Dictionary<string, List<Dictionary<string, string>>> dict = new Dictionary<string, List<Dictionary<string, string>>>();

                for (int row = 2; row <= rowCount; row++)
                {
                    string key = worksheet.Cells[row, 1].Value.ToString();
                    string subject = worksheet.Cells[row, 2].Value.ToString();
                    string credential = worksheet.Cells[row, 3].Value.ToString();
                    string teachField = worksheet.Cells[row, 4].Value.ToString();
                    string teachFieldName = worksheet.Cells[row, 5].Value.ToString();

                    Dictionary<string, string> subDict = new Dictionary<string, string>();
                    subDict.Add("Subject", subject);
                    subDict.Add("Credential", credential);
                    subDict.Add("TeachFieldName", teachFieldName);

                    if (!dict.ContainsKey(teachField))
                    {
                        dict.Add(teachField, new List<Dictionary<string, string>>());
                    }

                    dict[teachField].Add(subDict);
                }

                return dict;
            }
        }
    }
}

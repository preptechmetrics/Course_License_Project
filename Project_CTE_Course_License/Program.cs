using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace Project_CTE_Course_License
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\prept\Desktop\Git_projects\TB_C_Sharp\TB_Project\Project_CTE_Course_License\CTE_Codes_Project.xlsx";
            ExcelData excelData = new ExcelData();
            Dictionary<string, Dictionary<string, string>> dict = excelData.GetDictionary(filePath);

            int i = 0;
            foreach (var kvp in dict)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value["Subject"]}, {kvp.Value["Credential"]}, {kvp.Value["TeachField"]}, {kvp.Value["TeachFieldName"]}");

                i++;
                if (i >= 5) break;
            }

            Console.WriteLine("Please enter a Teaching Field code:");
            string teachFieldCode = Console.ReadLine();

            if (dict.ContainsKey(teachFieldCode))
            {
                Dictionary<string, string> subDict = dict[teachFieldCode];
                Console.WriteLine($"Subject: {subDict["Subject"]}");
                Console.WriteLine($"Credential: {subDict["Credential"]}");
            }
            else
            {
                Console.WriteLine("No data found for Teaching Field code: " + teachFieldCode);
            }

            Console.ReadKey();
        }
    }

    internal class ExcelData
    {
        public Dictionary<string, Dictionary<string, string>> GetDictionary(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                Dictionary<string, Dictionary<string, string>> dict = new Dictionary<string, Dictionary<string, string>>();

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
                    subDict.Add("TeachField", teachField);
                    subDict.Add("TeachFieldName", teachFieldName);

                    dict.Add(key, subDict);
                }

                return dict;



            }
        }
    }
}

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
            Dictionary<string, List<Dictionary<string, string>>> dict = excelData.GetDictionary(filePath);

            Console.WriteLine("Please enter a Teaching Field code:");
            string teachFieldCode = Console.ReadLine();

            if (dict.ContainsKey(teachFieldCode))
            {
                List<Dictionary<string, string>> subDictList = dict[teachFieldCode];
                foreach (var subDict in subDictList)
                {
                    Console.WriteLine($"Subject: {subDict["Subject"]}");
                    Console.WriteLine($"Credential: {subDict["Credential"]}");
                    Console.WriteLine($"TeachFieldName: {subDict["TeachFieldName"]}");
                    Console.WriteLine();
                }
            }
            else
            {
                Console.WriteLine("No data found for Teaching Field code: " + teachFieldCode);
            }

            Console.ReadKey();
        }
    }
}

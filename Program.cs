using System;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace ToExcelFromJsonDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonString = File.ReadAllText("Input/input.json");
            JsonDocument document = JsonDocument.Parse(jsonString);
            var _object = JsonSerializer.Deserialize<JsonElement>(jsonString);
            JsonElement.ObjectEnumerator ttst = _object.EnumerateObject();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet 1");
                var column = 1;
                var row = 1;
                var column2 = 1;
                foreach (var item in ttst)
                {
                    if (item.Value.ValueKind.ToString() != "Array"){
                        worksheet.Column(column).Cell(row).Value = item.Name;
                        worksheet.Column(column).Cell(row+1).Value = item.Value;
                    }
                    else{
                        column2 = column;
                        string firstCell = Convert.ToChar(column+64).ToString();
                        string lastCell = Convert.ToChar(column+item.Value.GetArrayLength()+63).ToString();
                        string rango = firstCell+"1:"+lastCell+"1";
                        worksheet.Range(rango).Merge();
                        worksheet.Column(column).Cell(row).Value = item.Name;
                        var test = item.Value.EnumerateArray();
                        foreach (var subitem in test)
                        {
                            worksheet.Column(column2).Cell(row+1).Value = subitem.ToString();
                            column2++;
                        }
                    }
                    column++;
                    row = row+2;
                }
                workbook.SaveAs("OutPut/Test.xlsx");
            }
            Console.WriteLine(_object.GetRawText());
            Console.Write(_object);
        }
    }
}

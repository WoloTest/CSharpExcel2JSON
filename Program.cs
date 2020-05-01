using System.IO;
using ExcelDataReader;
using Newtonsoft.Json;

namespace HelloWorld
{
    class Program
    {
        public const int MaxColumn = 7;
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var jsonFileStream = File.Open("json/test1.json", FileMode.Create))
            using (var jsonStreamWriter = new StreamWriter(jsonFileStream))
            using (var jsonWriter = new JsonTextWriter(jsonStreamWriter))
            {
                jsonWriter.Formatting = Formatting.Indented;
                jsonWriter.WriteStartArray();

                using (var excelFileStream = File.Open("xlsx/test1.xlsx", FileMode.Open, FileAccess.Read))
                {
                    using (var excelReader = ExcelReaderFactory.CreateReader(excelFileStream))
                    {
                        do
                        {
                            var row = 0;
                            while (excelReader.Read())
                            {
                                for (int column = 0; column < MaxColumn + 1; column++)
                                {
                                    jsonWriter.WriteStartObject();
                                    jsonWriter.WritePropertyName("column");
                                    jsonWriter.WriteValue(column);
                                    jsonWriter.WritePropertyName("row");
                                    jsonWriter.WriteValue(row);
                                    jsonWriter.WritePropertyName("type");
                                    jsonWriter.WriteValue(excelReader.GetFieldType(column)?.Name ?? null);
                                    jsonWriter.WritePropertyName("value");
                                    jsonWriter.WriteValue(excelReader.GetValue(column));
                                    jsonWriter.WriteEndObject();
                                }
                                row++;
                            }
                        } while (excelReader.NextResult());
                    }
                }

                jsonWriter.WriteEndArray();
            }
        }
    }
}

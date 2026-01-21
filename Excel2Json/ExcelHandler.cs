using System;
using System.Diagnostics;
using System.IO;
using Excel2Json.DocReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Excel2Json
{
    public class ExcelHandler
    {
        public static void Handle(string srcFileName, string dstFileName)
        {
            var fileInfor = new FileInfo(srcFileName);
            using (var package = new ExcelWorksheet(fileInfor)) {
                var sw = Stopwatch.StartNew();
                try {
                    var root = new ExcelHandler(package).Parse(srcFileName);
                    using (var wstream = File.CreateText(dstFileName)) {
                        using (var writer = new JsonTextWriter(wstream)) {
                            writer.Formatting = Formatting.Indented;
                            writer.IndentChar = '\t';
                            writer.Indentation = 1;
                            root.WriteTo(writer);
                        }
                        sw.Stop();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"Convert success {sw.Elapsed.TotalMilliseconds} ms: {Path.GetFileName(srcFileName)} --> {Path.GetFileName(dstFileName)}");
                    }
                }
                catch {
                    sw.Stop();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Convert fail : {Path.GetFileName(srcFileName)} --> {Path.GetFileName(dstFileName)}");
                    throw;
                }
            }
        }

        private readonly ExcelWorksheet worksheet;

        ExcelHandler(ExcelWorksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        private JToken Parse(string srcFileName)
        {
            var docType = worksheet.GetString(1, 1);
            var reader = MapReaderFactory.Create(docType);
            var tableHeader = TableHeader.Build(worksheet);
            return reader.Read(tableHeader, worksheet);
        }
    }
}

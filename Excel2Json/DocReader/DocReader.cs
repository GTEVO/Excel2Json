using Newtonsoft.Json.Linq;

namespace Excel2Json.DocReader
{
    public abstract class DocReader
    {
        public abstract JToken Read(TableHeader tableHeader, ExcelWorksheet worksheet);

        internal static int FindStartRow(ExcelWorksheet worksheet, int startRow)
        {
            var maxRow = worksheet.EndRow;
            do {
                ++startRow;
                var text = worksheet.GetString(startRow, 1);
                if (text is null) {
                    continue;
                }

                text = text.Trim().ToLower();
                if (text == "value") {
                    break;
                }
                else if (text == "#") {
                    startRow = -1;
                    break;
                }
            } while (startRow <= maxRow);

            return startRow;
        }
    }
}

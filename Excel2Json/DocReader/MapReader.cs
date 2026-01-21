using System.Diagnostics;
using Excel2Json.Node;
using Newtonsoft.Json.Linq;

namespace Excel2Json.DocReader
{
    internal class MapReader : DocReader
    {
        public override JToken Read(TableHeader tableHeader, ExcelWorksheet worksheet)
        {
            var result = new JObject();

            var nodeTree = ObjectNode.Create(tableHeader, worksheet);
            var startRow = FindStartRow(worksheet, tableHeader.EndRow);
            if (startRow == -1) {
                return result;
            }

            var endRow = startRow;
            var maxRow = worksheet.EndRow;
            do {
                ++endRow;
                var text = worksheet.GetString(endRow, 1);
                if (text is null) {
                    continue;
                }

                text = text.Trim().ToLower();

                if (text == "value" || text == "#") {
                    var key = worksheet.GetString(startRow, 2);
                    var (jToken, _) = nodeTree.Read(worksheet, startRow, endRow - 1);                    
                    result.Add(key, jToken);
                    startRow = endRow;
                }

                if (text == "#") {
                    break;
                }

            } while (endRow <= maxRow);

            return result;
        }
    }
}

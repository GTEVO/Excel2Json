using Excel2Json.Node;
using Newtonsoft.Json.Linq;

namespace Excel2Json.DocReader
{
    internal class ObjectReader : DocReader
    {
        public override JToken Read(TableHeader tableHeader, ExcelWorksheet worksheet)
        {
            var nodeTree = ObjectNode.Create(tableHeader, worksheet);
            var startRow = FindStartRow(worksheet, tableHeader.EndRow);
            if (startRow == -1) {
                return new JObject();
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
                if (text == "#") {
                    var (jToken, _) = nodeTree.Read(worksheet, startRow, endRow - 1);                    
                    return jToken;
                }

            } while (endRow <= maxRow);

            return new JObject();
        }
    }
}
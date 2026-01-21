using Excel2Json.Node;
using Newtonsoft.Json.Linq;

namespace Excel2Json.DocReader
{
    internal class ArrayReader : DocReader
    {
        public override JToken Read(TableHeader tableHeader, ExcelWorksheet worksheet)
        {
            var result = new JArray();

            var nodeTree = ArrayNode.Create(tableHeader, worksheet);
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
                    var (jToken, _) = nodeTree.Read(worksheet, startRow, endRow - 1);
                    foreach (var child in jToken.Children()) {
                        result.Add(child);
                    }
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

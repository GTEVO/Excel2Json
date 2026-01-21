using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class StrPropertyNode : PropertyNode
    {
        public StrPropertyNode(int layer, TableHeader tableHeader, int varRow, int col) : base(layer, tableHeader, varRow, col)
        {
        }

        protected override JToken Read(ExcelWorksheet worksheet, int row)
        {
            var value = worksheet.GetString(row, StartCol);
            return new JValue(value);
        }
    }
}
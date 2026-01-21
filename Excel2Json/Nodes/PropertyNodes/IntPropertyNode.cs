using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class IntPropertyNode : PropertyNode
    {
        public IntPropertyNode(int layer, TableHeader tableHeader, int varRow, int col) : base(layer, tableHeader, varRow, col)
        {
        }

        protected override JToken Read(ExcelWorksheet worksheet, int row)
        {
            var value = worksheet.GetValue<int>(row, StartCol);
            return new JValue(value);
        }
    }
}
using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class BooleanPropertyNode : PropertyNode
    {
        public BooleanPropertyNode(int layer, TableHeader tableHeader, int varRow, int col) : base(layer, tableHeader, varRow, col)
        {
        }

        protected override JToken Read(ExcelWorksheet worksheet, int row)
        {
            var value = worksheet.GetValue<bool>(row, StartCol);
            return new JValue(value);
        }
    }
}
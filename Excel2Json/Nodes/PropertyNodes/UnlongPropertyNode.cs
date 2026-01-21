using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class UnlongPropertyNode : PropertyNode
    {
        public UnlongPropertyNode(int layer, TableHeader tableHeader, int varRow, int col) : base(layer, tableHeader, varRow, col)
        {
        }

        protected override JToken Read(ExcelWorksheet worksheet, int row)
        {
            var value = worksheet.GetValue<ulong>(row, StartCol);
            return new JValue(value);
        }
    }
}

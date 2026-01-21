using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal abstract class PropertyNode : Node
    {
        public PropertyNode(int layer, TableHeader tableHeader, int varRow, int col)
        {
            string varName = null;
            if (varRow != -1) {
                varName = tableHeader.GetTextSafe(varRow, col);
            }

            VarName = varName;
            StartCol = col;
            EndCol = col;
            Type = NodeType.Property;
            Layer = layer;
        }

        protected abstract JToken Read(ExcelWorksheet worksheet, int row);

        public override void Build(TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude)
        {

        }

        public override void Update(Node parent)
        {

        }

        public override (JToken jToken, int hasReadRow) Read(ExcelWorksheet worksheet, int startRowInclude, int endRowInclude)
        {
            var value = Read(worksheet, startRowInclude);
            return (value, startRowInclude);
        }
    }
}
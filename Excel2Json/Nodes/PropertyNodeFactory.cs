using System;

namespace Excel2Json.Node
{
    internal class PropertyNodeFactory
    {
        public static Node Create(string type, int layer, TableHeader tableHeader, ExcelWorksheet worksheet, int varRow, int col, int maxCol)
        {
            switch (type) {
                case "i":
                    return new IntPropertyNode(layer, tableHeader, varRow, col);
                case "f":
                    return new FloatPropertyNode(layer, tableHeader, varRow, col);
                case "s":
                    return new StrPropertyNode(layer, tableHeader, varRow, col);
                case "b":
                    return new BooleanPropertyNode(layer, tableHeader, varRow, col);
                case "ul":
                    return new UnlongPropertyNode(layer, tableHeader, varRow, col);
                case "[]":
                    return ArrayNode.Create(layer, tableHeader, varRow, col, maxCol);
                case "{}":
                    return ObjectNode.Create(layer, tableHeader, varRow, col, maxCol);
                default:
                    throw new NotSupportedException($"Not Supported type: {type}");
            }
        }
    }
}
using System;
using System.Collections.Generic;

namespace Excel2Json.Node
{
    internal static class Utils
    {
        // 从varRow开始，从startCol到endCol结束，构建nodes
        public static void BuildWithVarRow(Node node, TableHeader tableHeader, ExcelWorksheet worksheet, int varRow, int startCol, int endCol)
        {
            var children = node.Children;

            for (int i = startCol; i <= endCol;) {
                //  找到描述varRow的typeRow
                var (typeRow, type) = tableHeader.GetVarRowTypeRow(varRow, i);
                if (typeRow == -1) {
                    throw new Exception($"[{varRow},{startCol}] not found header");
                }
                var propertyNode = PropertyNodeFactory.Create(type, node.Layer + 1, tableHeader, worksheet, varRow, i, endCol);
                i = propertyNode.EndCol + 1;
                propertyNode.Build(tableHeader, worksheet, typeRow);
                children.Add(propertyNode);
            }
        }

    }

}

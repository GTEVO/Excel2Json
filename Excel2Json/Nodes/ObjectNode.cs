using System;
using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class ObjectNode : Node
    {
        public static ObjectNode Create(TableHeader tableHeader, ExcelWorksheet worksheet)
        {
            var node = Create(0, tableHeader, worksheet, 1, tableHeader.StartCol, tableHeader.EndCol);
            node.Update(null);
            return node;
        }

        //  用于创建没有属性名的object节点，例如数组的元素对象、字典类型文档的value对象
        public static ObjectNode Create(int layer, TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude, int startColInclude, int endColInclude)
        {
            var node = new ObjectNode {
                StartCol = startColInclude,
                EndCol = endColInclude,
                Children = [],
                Type = NodeType.Object,
                Layer = layer,
            };
            node.Build(tableHeader, worksheet, startRowExclude);
            return node;
        }

        public static ObjectNode Create(int layer, TableHeader tableHeader, int varRow, int col, int maxCol)
        {
            var varName = tableHeader.GetTextSafe(varRow, col);
            int endCol = tableHeader.GetObjOrArrayNodeEndCol(varRow, col, maxCol);

            var node = new ObjectNode() {
                VarName = varName,
                StartCol = col,
                EndCol = endCol,
                Children = [],
                Type = NodeType.Object,
                Layer = layer,
            };
            return node;
        }

        public override void Build(TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude)
        {
            var row = startRowExclude + 1;

            var rowType = tableHeader.GetRowType(row);
            switch (rowType) {
                case TableHeaderType.Var:
                    Utils.BuildWithVarRow(this, tableHeader, worksheet, row, StartCol, EndCol);
                    break;
                default:
                    // 去到下一行找var
                    Utils.BuildWithVarRow(this, tableHeader, worksheet, row + 1, StartCol, EndCol);
                    break;
            }
        }

        public override void Update(Node parent)
        {
            foreach (var child in Children) {
                if (child.Type == NodeType.Array) {
                    child.Update(this);
                }
            }
        }

        public override (JToken jToken, int hasReadRow) Read(ExcelWorksheet worksheet, int startRowInclude, int endRowInclude)
        {
            var __hasReadRow = startRowInclude;
            JObject jObject = null;

            for (int col = StartCol; col <= EndCol; col++) {
                if (worksheet.HasValue(startRowInclude, col)) {
                    goto read;
                }
            }
            goto end;

        read:
            jObject = new();
            foreach (var child in Children) {
                var (value, hasReadRow) = child.Read(worksheet, startRowInclude, endRowInclude);
                jObject.Add(child.VarName, value);
                __hasReadRow = Math.Max(__hasReadRow, hasReadRow);
            }

        end:
            return (jObject, __hasReadRow);
        }
    }
}
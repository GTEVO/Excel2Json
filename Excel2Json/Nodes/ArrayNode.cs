using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    internal class ArrayNode : Node
    {
        private readonly List<Node> _baseLineCheckNodes = new();

        public static ArrayNode Create(TableHeader tableHeader, ExcelWorksheet worksheet)
        {
            var node = Create(0, tableHeader, worksheet, 1, tableHeader.StartCol, tableHeader.EndCol);
            node.Update(null);
            return node;
        }

        //  用于创建没有属性名的array节点，例如数组类型文档的elment对象
        public static ArrayNode Create(int layer, TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude, int startColInclude, int endColInclude)
        {
            var node = new ArrayNode {
                StartCol = startColInclude,
                EndCol = endColInclude,
                Children = [],
                Type = NodeType.Array,
                Layer = layer,
            };
            node.Build(tableHeader, worksheet, startRowExclude);
            return node;
        }

        public static Node Create(int layer, TableHeader tableHeader, int varRow, int col, int maxCol)
        {
            var varName = tableHeader.GetTextSafe(varRow, col);
            int endCol = tableHeader.GetObjOrArrayNodeEndCol(varRow, col, maxCol);

            var node = new ArrayNode() {
                VarName = varName,
                StartCol = col,
                EndCol = endCol,
                Type = NodeType.Array,
                Layer = layer,
            };
            return node;
        }

        public override void Build(TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude)
        {
            for (int row = startRowExclude + 1; ; row++) {
                var type = tableHeader.GetRowType(row);
                switch (type) {
                    case TableHeaderType.Type: {
                            //  离var最近的Type行，必定是元素类型
                            var elementType = tableHeader.GetTextUnSafe(row, StartCol);
                            if (string.IsNullOrEmpty(elementType)) {
                                continue;
                            }
                            Children = [];
                            var propertyNode = PropertyNodeFactory.Create(elementType, Layer + 1, tableHeader, worksheet, -1, StartCol, EndCol);
                            if (propertyNode.Type == NodeType.Array) {
                                throw new NotSupportedException("not support 2D array !");
                            }
                            propertyNode.Build(tableHeader, worksheet, row);
                            Children.Add(propertyNode);
                        }
                        return;
                    case TableHeaderType.Var: {
                            var text = tableHeader.GetTextUnSafe(row, StartCol);
                            if (!string.IsNullOrEmpty(text)) {
                                //  离var最近的是Var行，则元素类型是object
                                Children = [];
                                var objNode = ObjectNode.Create(Layer + 1, tableHeader, worksheet, row - 1, StartCol, EndCol);
                                Children.Add(objNode);
                                return;
                            }
                        }
                        break;
                }
            }
        }

        public override void Update(Node parent)
        {
            static void findBaseLineCheckNodes(Node bro, List<Node> results)
            {
                foreach (var child in bro.Children) {
                    switch (child.Type) {
                        case NodeType.Property: {
                                results.Add(child);
                            }
                            break;
                        case NodeType.Object: {
                                findBaseLineCheckNodes(child, results);
                            }
                            break;
                    }
                }
            }

            if (Layer > 1) {
                findBaseLineCheckNodes(parent, _baseLineCheckNodes);
            }

            foreach (var child in Children) {
                child.Update(this);
            }
        }

        public override (JToken jToken, int hasReadRow) Read(ExcelWorksheet worksheet, int startRowInclude, int endRowInclude)
        {
            //  Excel行转化为复杂json对象的原理：
            //
            //  将json对象用树形结构表示，每个子节点占用一个单元格，将这个结构填到Excel中，并将节点名和节点类型各行排列，即为该json结构配置表的表头；
            //  每个叶子节点对应一列，紧凑排列成一行。读取值时，按行读取，将单元格中的值按照json对象的树形结构进行还原，即可得到目标结构的json对象。
            //
            //  一、当json结构不包含数组时，仅用一行则可以表示一个json对象的值，我们将这行定义为【基础行】。
            //  二、当json结构包含数组时，由于每个数组元素对象占用一行，因此当存在数组节点时，数组有会将父节点占用的行数提高为数组元素的个数，从而可能超过一行。
            //  我们将数组元素占用的非基础行定义为【扩展行】。
            //      a、当存在多个但不嵌套的数组结构时，一个节点占用的行数等于子节点数组中元素数量最多的那个数组的元素数量。
            //      b、当存在嵌套数组时，节点占用行数的计算会更加复杂，因为下层数组的扩展行行数会影响上层数组的元素的扩展行的位置，
            //      因此读取嵌套数组对象时，需要从最深数组节点开始读取，然后逐层向上，直到最上层的节点。
            //      c、数组结束的判断: 
            //          1、用于表示数组元素行上子节点对应的列全为空值。
            //          2、根据一、二的说明，当读取到基础行时，则数组读取完毕。在Excel中其特征为数组节点的兄弟节点行存在非空值，则该行已经超出数组读取范围，并将该行号
            //          返回给上层组数，上层数组立即读取一个元素，在读取下一个元素时，则从该行开始读取，重复该行为直到顶层节点读取完毕。
            //
            //  缺陷: 2D数组是没办法通过单元格内的值来（构建基础行，从而）判断第一层数组的开始和结束位置，需要额外的符号来标记，为了简化Excel的配置难度，直接放弃支持2D数组。

            if (Children?.Count == 1) {
                var type = Children[0].Type;
                switch (type) {
                    case NodeType.Object:
                        return ReadObjectArray(Children[0], worksheet, startRowInclude, endRowInclude);
                    case NodeType.Property:
                        return ReadObjectArray(Children[0], worksheet, startRowInclude, endRowInclude);
                    default:
                        throw new NotSupportedException($"not support array type : {type}");
                }
            }

            throw new NotSupportedException($"not support array type : [{StartCol},{EndCol}]");
        }

        private (JToken jToken, int hasReadRow) ReadObjectArray(Node node, ExcelWorksheet worksheet, int startRowInclude, int endRowInclude)
        {
            int __hasReadRow = startRowInclude;
            var jArray = new JArray();
            for (int row = startRowInclude; row <= endRowInclude; row++) {

                //  检查基础行单元格内是否有值
                //  注意：只在扩展行进行检查
                if (row > startRowInclude) {
                    foreach (var brother in _baseLineCheckNodes) {
                        if (worksheet.HasValue(row, brother.StartCol)) {
                            goto end;
                        }
                    }
                }

                //  检查元素节点范围内的单元格是否有值
                for (int col = StartCol; col <= EndCol; col++) {
                    if (worksheet.HasValue(row, col)) {
                        goto read;
                    }
                }
                goto end;

            read:
                var (value, hasReadRow) = node.Read(worksheet, row, endRowInclude);
                __hasReadRow = Math.Max(__hasReadRow, hasReadRow);
                jArray.Add(value);
                row = hasReadRow;
            }

        end:
            return (jArray, __hasReadRow);
        }
    }
}

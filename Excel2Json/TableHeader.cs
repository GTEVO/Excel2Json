using System;
using System.Collections.Generic;

namespace Excel2Json
{
    public enum TableHeaderType
    {
        None,
        Var,
        Type,
    }

    public class TableHeader
    {
        public ExcelWorksheet Worksheet { get; private set; }

        public int StartCol { get; }
        public int EndCol { get; }
        public int EndRow { get; }

        private readonly List<TableHeaderType> _tableHeaderTypes;


        TableHeader(int startCol, int endRow, int endCol, List<TableHeaderType> tableHeaderTypes)
        {
            StartCol = startCol;
            EndRow = endRow;
            EndCol = endCol;
            _tableHeaderTypes = tableHeaderTypes;
        }

        public static TableHeader Build(ExcelWorksheet worksheet)
        {
            int startCol = 0;
            int endRow = 0;
            int width = 0;
            var tableHeaderTypes = new List<TableHeaderType>() { TableHeaderType.None };

            var maxColumn = worksheet.EndColumn;
            for (int col = 1; col <= maxColumn; col++) {
                var text = worksheet.GetString(1, col);
                if (text == "$") {
                    startCol = col;
                    break;
                }
            }

            for (int col = 1; col <= maxColumn; col++) {
                var text = worksheet.GetString(1, col);
                if (text == "#") {
                    width = col;
                    break;
                }
            }

            var maxRow = worksheet.EndRow;
            for (int row = 2; row <= maxRow; row++) {
                var text = worksheet.GetString(row, 1);
                switch (text) {
                    case "##var":
                        tableHeaderTypes.Add(TableHeaderType.Var);
                        break;
                    case "##type":
                        tableHeaderTypes.Add(TableHeaderType.Type);
                        break;
                    default:
                        tableHeaderTypes.Add(TableHeaderType.None);
                        break;
                }

                if (worksheet.GetString(row, width) == "#") {
                    endRow = row;
                    break;
                }
            }

            if (startCol == 0) {
                throw new Exception("table header startCol not found");
            }
            if (width == 0) {
                throw new Exception("table header endCol not found");
            }
            if (endRow == 0) {
                throw new Exception("table header endRow not found");
            }

            return new TableHeader(startCol, endRow, width - 1, tableHeaderTypes) {
                Worksheet = worksheet,
            };
        }

        public string GetTextSafe(int row, int col)
        {
            var varName = GetTextUnSafe(row, col);
            if (string.IsNullOrEmpty(varName)) {
                throw new ArgumentNullException($"[{row},{col}] is empty var");
            }
            return varName;
        }

        public string GetTextUnSafe(int row, int col)
        {
            var varName = Worksheet.GetString(row, col);
            if (string.IsNullOrEmpty(varName))
                return varName;
            return varName.Trim();
        }

        public TableHeaderType GetRowType(int row)
        {
            return _tableHeaderTypes[row - 1];
        }

        /// <summary>
        /// varRow为Array或Object的var行
        /// <returns></returns>
        public int GetObjOrArrayNodeEndCol(int varRow, int startColExclude, int maxCol)
        {
            int endCol = startColExclude;
            while (true) {
                var nextCol = endCol + 1;
                if (nextCol > EndCol || nextCol > maxCol) {
                    break;
                }

                var nextVar = Worksheet.GetString(varRow, nextCol);
                if (!string.IsNullOrEmpty(nextVar)) {
                    break;
                }
                endCol = nextCol;
            }

            return endCol;
        }

        public (int row, string type) GetVarRowTypeRow(int varRow, int col)
        {
            var index = varRow - 1;

            //  从下一行开始查找
            index += 1;

            while (index < _tableHeaderTypes.Count) {
                if (_tableHeaderTypes[index] == TableHeaderType.Type) {
                    // Row 从 1 开始，index 从 0 开始 ，所以需要加 1
                    var row = index + 1;
                    var type = Worksheet.GetString(row, col);
                    if (!string.IsNullOrEmpty(type)) {
                        return (row, type);
                    }
                }
                ++index;
            }

            return (-1, null);
        }

    }
}

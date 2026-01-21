using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Excel2Json.Node
{
    public enum NodeType
    {
        Object,
        Array,
        Property,
    }

    public abstract class Node
    {
        public string VarName { get; protected set; }
        public int Layer { get; protected set; }

        public int StartCol { get; protected set; }
        public int EndCol { get; protected set; }

        public NodeType Type { get; protected set; } = NodeType.Property;
        public List<Node> Children { get; protected set; }

        public abstract void Build(TableHeader tableHeader, ExcelWorksheet worksheet, int startRowExclude);
        public abstract void Update(Node parent);
        public abstract (JToken jToken, int hasReadRow) Read(ExcelWorksheet worksheet, int startRowInclude, int endRowInclude);
    }
}